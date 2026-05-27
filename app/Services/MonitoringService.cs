namespace MailMonitor.Services;

/// <summary>
/// Orchestrateur monitoring : reconcile config → initial scan par dossier
/// (avec watermark `last_received_ts`) → boucle de polling périodique (30s)
/// qui ne récupère que les mails reçus depuis le dernier passage.
/// Pas de COM events → plus simple, plus robuste, latence ~30s acceptable.
/// </summary>
public sealed class MonitoringService
{
    private readonly OutlookService _outlook;
    private readonly StorageService _storage;
    private readonly ClassificationService _classifier;
    private readonly LogService _log;

    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, long> _folderDbIds = new();
    private CancellationTokenSource? _cts;
    private Task? _pollTask;
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(30);

    public event Action<OutlookEvent>? OnMailEvent;
    public event Action? OnStatsChanged;

    public bool IsRunning { get; private set; }
    public OutlookService Outlook => _outlook;

    public MonitoringService(OutlookService outlook, StorageService storage, ClassificationService cls, LogService log)
    {
        _outlook = outlook; _storage = storage; _classifier = cls; _log = log;
    }

    public async Task StartAsync()
    {
        if (IsRunning) return;
        _cts = new CancellationTokenSource();
        IsRunning = true;
        _log.Info("MONITOR", "Démarrage…");

        try { await _outlook.WhenReady; }
        catch (Exception ex)
        {
            _log.Warn("MONITOR", "Outlook indisponible — monitoring inactif: " + ex.Message);
            IsRunning = false;
            return;
        }

        await ReconcileFoldersAsync();
        await PollOnceAsync(_cts.Token, isInitial: true);

        _pollTask = Task.Run(() => PollLoopAsync(_cts.Token));

        _log.Info("MONITOR", $"Opérationnel — polling toutes les {PollInterval.TotalSeconds:F0}s.");
        OnStatsChanged?.Invoke();
    }

    public void Stop()
    {
        IsRunning = false;
        try { _cts?.Cancel(); } catch { }
    }

    private async Task PollLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try { await Task.Delay(PollInterval, ct); }
            catch (OperationCanceledException) { return; }

            try { await PollOnceAsync(ct, isInitial: false); }
            catch (Exception ex) { _log.Warn("POLL", "Cycle KO: " + ex.Message); }
        }
    }

    private async Task ReconcileFoldersAsync()
    {
        _folderDbIds.Clear();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT id, entry_id FROM folders WHERE is_monitored=1";
        using var r = cmd.ExecuteReader();
        while (r.Read()) _folderDbIds[r.GetString(1)] = r.GetInt64(0);
        _log.Info("MONITOR", $"{_folderDbIds.Count} dossier(s) monitoré(s) en config");
        await Task.CompletedTask;
    }

    private async Task PollOnceAsync(CancellationToken ct, bool isInitial)
    {
        if (_folderDbIds.Count == 0) return;

        var meta = new List<(long id, string storeId, string entryId, string category, long? since)>();
        using (var c = _storage.OpenConnection())
        using (var cmd = c.CreateCommand())
        {
            cmd.CommandText = "SELECT id, store_id, entry_id, category, last_received_ts FROM folders WHERE is_monitored=1";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                meta.Add((r.GetInt64(0), r.GetString(1), r.GetString(2), r.GetString(3),
                    r.IsDBNull(4) ? null : r.GetInt64(4)));
        }

        int totalNew = 0;
        foreach (var f in meta)
        {
            if (ct.IsCancellationRequested) break;
            try
            {
                DateTime? since = f.since.HasValue
                    ? DateTimeOffset.FromUnixTimeSeconds(f.since.Value).LocalDateTime
                    : null;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var mails = await _outlook.ScanFolderAsync(f.storeId, f.entryId, since);
                if (mails.Count == 0)
                {
                    // Mise à jour last_scan_ts (+ init last_received_ts si NULL) pour éviter
                    // un rescan complet à chaque cycle sur dossiers vides (review Copilot).
                    TouchFolderScan(f.id);
                    continue;
                }
                _classifier.BulkUpsertEmails(f.id, f.category, mails);
                long maxTs = mails.Max(m => new DateTimeOffset(m.ReceivedTime).ToUnixTimeSeconds());
                UpdateFolderWatermark(f.id, maxTs);
                totalNew += mails.Count;

                foreach (var m in mails)
                {
                    OnMailEvent?.Invoke(new OutlookEvent
                    {
                        Type = OutlookEvent.Kind.Add,
                        StoreId = f.storeId,
                        FolderEntryId = f.entryId,
                        MailEntryId = m.EntryId,
                        Mail = m
                    });
                }

                if (isInitial)
                    _log.Info("SCAN", $"{f.entryId[..Math.Min(8, f.entryId.Length)]}…  +{mails.Count} mails en {sw.ElapsedMilliseconds}ms");
            }
            catch (Exception ex) { _log.Error("SCAN", "Échec dossier " + f.entryId, ex); }
        }

        if (totalNew > 0)
        {
            _log.Info("POLL", $"+{totalNew} mails (delta)");
            OnStatsChanged?.Invoke();
        }
        else if (isInitial) OnStatsChanged?.Invoke();
    }

    private void UpdateFolderWatermark(long folderId, long lastReceivedTs)
    {
        lock (_storage.WriteLock)
        {
            using var c = _storage.OpenConnection();
            using var cmd = c.CreateCommand();
            cmd.CommandText = "UPDATE folders SET last_scan_ts=$now, last_received_ts=$ts WHERE id=$id";
            cmd.Parameters.AddWithValue("$now", DateTimeOffset.UtcNow.ToUnixTimeSeconds());
            cmd.Parameters.AddWithValue("$ts", lastReceivedTs);
            cmd.Parameters.AddWithValue("$id", folderId);
            cmd.ExecuteNonQuery();
        }
    }

    // Met à jour uniquement last_scan_ts, et initialise last_received_ts à maintenant
    // s'il est NULL (évite rescans complets sur dossiers vides). Review Copilot.
    private void TouchFolderScan(long folderId)
    {
        lock (_storage.WriteLock)
        {
            using var c = _storage.OpenConnection();
            using var cmd = c.CreateCommand();
            cmd.CommandText = @"UPDATE folders
               SET last_scan_ts=$now,
                   last_received_ts=COALESCE(last_received_ts, $now)
               WHERE id=$id";
            cmd.Parameters.AddWithValue("$now", DateTimeOffset.UtcNow.ToUnixTimeSeconds());
            cmd.Parameters.AddWithValue("$id", folderId);
            cmd.ExecuteNonQuery();
        }
    }

    public async Task AddFolderAsync(string storeId, string entryId, string path, string displayName, string? category = null)
    {
        category ??= ClassificationService.ClassifyByPath(path);
        long id;
        lock (_storage.WriteLock)
        {
            using var c = _storage.OpenConnection();
            using var cmd = c.CreateCommand();
            cmd.CommandText = @"
INSERT INTO folders(store_id, entry_id, path, display_name, category, is_monitored)
VALUES($s,$e,$p,$d,$c,1)
ON CONFLICT(store_id, entry_id) DO UPDATE SET path=excluded.path, display_name=excluded.display_name,
   category=excluded.category, is_monitored=1
RETURNING id;";
            cmd.Parameters.AddWithValue("$s", storeId);
            cmd.Parameters.AddWithValue("$e", entryId);
            cmd.Parameters.AddWithValue("$p", path);
            cmd.Parameters.AddWithValue("$d", displayName);
            cmd.Parameters.AddWithValue("$c", category);
            id = (long)cmd.ExecuteScalar()!;
        }
        _folderDbIds[entryId] = id;
        var mails = await _outlook.ScanFolderAsync(storeId, entryId, null);
        _classifier.BulkUpsertEmails(id, category, mails);
        if (mails.Count > 0)
        {
            long maxTs = mails.Max(m => new DateTimeOffset(m.ReceivedTime).ToUnixTimeSeconds());
            UpdateFolderWatermark(id, maxTs);
        }
        else
        {
            // Init watermark à maintenant sinon rescan complet à chaque polling (review Copilot)
            TouchFolderScan(id);
        }
        OnStatsChanged?.Invoke();
    }

    public Task RemoveFolderAsync(string entryId)
    {
        lock (_storage.WriteLock)
        {
            using var c = _storage.OpenConnection();
            using var cmd = c.CreateCommand();
            cmd.CommandText = "UPDATE folders SET is_monitored=0 WHERE entry_id=$e";
            cmd.Parameters.AddWithValue("$e", entryId);
            cmd.ExecuteNonQuery();
        }
        _folderDbIds.TryRemove(entryId, out _);
        OnStatsChanged?.Invoke();
        return Task.CompletedTask;
    }
}
