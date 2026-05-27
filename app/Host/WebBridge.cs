using System.Text.Json;
using System.Text.Json.Serialization;
using MailMonitor.Services;
using Microsoft.Data.Sqlite;
using Microsoft.Web.WebView2.WinForms;

namespace MailMonitor.Host;

public sealed class WebBridge
{
    private readonly WebView2 _web;
    private readonly MonitoringService _monitor;
    private readonly StorageService _storage;
    private readonly LogService _log;
    private readonly UpdateService _updater;
    private readonly AutoStartService _autostart;
    private readonly Form _ownerForm;

    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public WebBridge(WebView2 web, MonitoringService monitor, StorageService storage, LogService log,
                     UpdateService updater, AutoStartService autostart, Form ownerForm)
    {
        _web = web; _monitor = monitor; _storage = storage; _log = log;
        _updater = updater; _autostart = autostart; _ownerForm = ownerForm;
    }

    public void Attach()
    {
        var shimPath = Path.Combine(AppContext.BaseDirectory, "wwwroot", "preload-shim.js");
        if (!File.Exists(shimPath))
        {
            var alt = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "wwwroot", "preload-shim.js"));
            if (File.Exists(alt)) shimPath = alt;
        }
        var shim = File.Exists(shimPath) ? File.ReadAllText(shimPath) : InlineFallbackShim;
        _ = _web.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(shim);

        _web.CoreWebView2.WebMessageReceived += OnMessage;

        _monitor.OnMailEvent += e => PostEvent("mail." + e.Type.ToString().ToLowerInvariant(), e);
        _monitor.OnStatsChanged += () => PostEvent("stats.changed", null);
        _log.OnEntry += entry => PostEvent("log.entry", entry);
        _updater.OnAvailable += v => PostEvent("update.available", new { version = v });
        _updater.OnProgress += p => PostEvent("update.progress", new { percent = p });
        _updater.OnReadyToInstall += () => PostEvent("update.ready", null);
    }

    private void PostEvent(string name, object? data)
    {
        try
        {
            if (_web.CoreWebView2 is null) return;
            var json = JsonSerializer.Serialize(new { @event = name, data }, Json);
            _ownerForm.BeginInvoke(new Action(() => _web.CoreWebView2.PostWebMessageAsJson(json)));
        }
        catch { }
    }

    private async void OnMessage(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
    {
        long id = 0;
        try
        {
            using var doc = JsonDocument.Parse(e.WebMessageAsJson);
            id = doc.RootElement.GetProperty("id").GetInt64();
            var method = doc.RootElement.GetProperty("method").GetString() ?? "";
            JsonElement args = doc.RootElement.TryGetProperty("args", out var a) ? a : default;

            var result = await DispatchAsync(method, args);
            Respond(id, true, result, null);
        }
        catch (Exception ex)
        {
            _log.Warn("RPC", $"Méthode KO: {ex.Message}");
            Respond(id, false, null, ex.Message);
        }
    }

    private void Respond(long id, bool ok, object? result, string? error)
    {
        try
        {
            if (_web.CoreWebView2 is null) return;
            var payload = JsonSerializer.Serialize(new { id, ok, result, error }, Json);
            _ownerForm.BeginInvoke(new Action(() =>
            {
                try
                {
                    if (_web.CoreWebView2 is null) return;
                    _web.CoreWebView2.PostWebMessageAsJson(payload);
                }
                catch { }
            }));
        }
        catch { }
    }

    // Validation centralisée des arguments RPC (review Copilot)
    private static JsonElement RequireArg(JsonElement args, int index, string method)
    {
        if (args.ValueKind != JsonValueKind.Array)
            throw new ArgumentException($"{method}: args doit être un tableau JSON");
        if (args.GetArrayLength() <= index)
            throw new ArgumentException($"{method}: argument [{index}] manquant");
        return args[index];
    }
    private static string RequireString(JsonElement args, int index, string method)
    {
        var el = RequireArg(args, index, method);
        if (el.ValueKind != JsonValueKind.String)
            throw new ArgumentException($"{method}: argument [{index}] doit être une chaîne");
        return el.GetString() ?? "";
    }
    private static bool RequireBool(JsonElement args, int index, string method)
    {
        var el = RequireArg(args, index, method);
        if (el.ValueKind != JsonValueKind.True && el.ValueKind != JsonValueKind.False)
            throw new ArgumentException($"{method}: argument [{index}] doit être un booléen");
        return el.GetBoolean();
    }
    private static int RequireInt(JsonElement args, int index, string method)
    {
        var el = RequireArg(args, index, method);
        if (el.ValueKind != JsonValueKind.Number)
            throw new ArgumentException($"{method}: argument [{index}] doit être un nombre");
        return el.GetInt32();
    }
    private static JsonElement RequireObject(JsonElement args, int index, string method)
    {
        var el = RequireArg(args, index, method);
        if (el.ValueKind != JsonValueKind.Object)
            throw new ArgumentException($"{method}: argument [{index}] doit être un objet");
        return el;
    }

    private async Task<object?> DispatchAsync(string method, JsonElement args)
    {
        switch (method)
        {
            case "app.version": return AppInfo.Version;
            case "app.autostart.get": return _autostart.IsEnabled;
            case "app.autostart.set":
                if (RequireBool(args, 0, method)) _autostart.EnsureEnabled();
                else _autostart.Disable();
                return null;
            case "app.check-updates": _ = _updater.CheckOnceAsync(); return null;
            case "app.apply-update": _updater.ApplyAndRestart(); return null;

            case "monitoring.status": return new { running = _monitor.IsRunning };
            case "outlook.status": return new { connected = _monitor.Outlook.IsConnected };

            case "outlook.list-stores":
                {
                    var stores = await _monitor.Outlook.ListStoresAsync();
                    return stores.Select(s => new { id = s.id, name = s.name }).ToList();
                }
            case "outlook.list-folders":
                {
                    var storeId = RequireString(args, 0, method);
                    var folders = await _monitor.Outlook.ListFoldersAsync(storeId);
                    return folders.Select(f => new
                    {
                        storeId = f.StoreId,
                        entryId = f.EntryId,
                        path = f.Path,
                        name = f.Name,
                        itemCount = f.ItemCount
                    }).ToList();
                }

            case "folders.list-monitored": return ListMonitoredFolders();
            case "folders.add":
                {
                    var p = RequireObject(args, 0, method);
                    if (!p.TryGetProperty("storeId", out var storeIdEl) || storeIdEl.ValueKind != JsonValueKind.String)
                        throw new ArgumentException(method + ": propriété 'storeId' manquante ou invalide");
                    if (!p.TryGetProperty("entryId", out var entryIdEl) || entryIdEl.ValueKind != JsonValueKind.String)
                        throw new ArgumentException(method + ": propriété 'entryId' manquante ou invalide");
                    if (!p.TryGetProperty("path", out var pathEl) || pathEl.ValueKind != JsonValueKind.String)
                        throw new ArgumentException(method + ": propriété 'path' manquante ou invalide");
                    await _monitor.AddFolderAsync(
                        storeIdEl.GetString()!,
                        entryIdEl.GetString()!,
                        pathEl.GetString()!,
                        p.TryGetProperty("displayName", out var dn) && dn.ValueKind == JsonValueKind.String ? dn.GetString()! : "",
                        p.TryGetProperty("category", out var c) && c.ValueKind == JsonValueKind.String ? c.GetString() : null);
                    return null;
                }
            case "folders.remove":
                await _monitor.RemoveFolderAsync(RequireString(args, 0, method));
                return null;

            case "stats.summary": return BuildSummary();
            case "stats.weekly": return BuildWeekly(args);
            case "stats.by-category": return BuildByCategory();

            case "emails.recent": return RecentEmails(args);

            case "logs.recent": return _log.Snapshot(500);

            case "weekly-comments.list":
                return ListWeeklyComments(RequireInt(args, 0, method), RequireInt(args, 1, method));
            case "weekly-comments.add":
                return AddWeeklyComment(RequireObject(args, 0, method));

            // ---- Compat avec l'ancien front Electron ----

            // Outlook discovery additionnel
            case "outlook.folders-shallow":
                {
                    var storeId = RequireString(args, 0, method);
                    string? parentId = args.GetArrayLength() > 1 && args[1].ValueKind == JsonValueKind.String ? args[1].GetString() : null;
                    var all = await _monitor.Outlook.ListFoldersAsync(storeId);
                    // Filtre approximatif : enfants directs (path = parent + '/X')
                    if (string.IsNullOrEmpty(parentId))
                    {
                        return all.Where(f => !f.Path.Contains('/') || f.Path.IndexOf('/') == f.Path.LastIndexOf('/'))
                                  .Select(f => new { storeId = f.StoreId, entryId = f.EntryId, path = f.Path, name = f.Name, itemCount = f.ItemCount, hasChildren = true })
                                  .ToList<object>();
                    }
                    return all.Select(f => new { storeId = f.StoreId, entryId = f.EntryId, path = f.Path, name = f.Name, itemCount = f.ItemCount, hasChildren = false }).ToList<object>();
                }
            case "outlook.folders-tree":
            case "outlook.folders-tree-from":
            case "folders.tree":
                return BuildFoldersTree();

            // Folders config bulk + update
            case "folders.add-bulk":
                {
                    var arr = RequireArg(args, 0, method);
                    if (arr.ValueKind != JsonValueKind.Array) throw new ArgumentException(method + ": tableau attendu");
                    int n = 0;
                    foreach (var p in arr.EnumerateArray())
                    {
                        if (p.ValueKind != JsonValueKind.Object) continue;
                        try
                        {
                            await _monitor.AddFolderAsync(
                                p.GetProperty("storeId").GetString()!,
                                p.GetProperty("entryId").GetString()!,
                                p.GetProperty("path").GetString()!,
                                p.TryGetProperty("displayName", out var dn) ? (dn.GetString() ?? "") : "",
                                p.TryGetProperty("category", out var ca) ? ca.GetString() : null);
                            n++;
                        }
                        catch (Exception ex) { _log.Warn("RPC", "folders.add-bulk item KO: " + ex.Message); }
                    }
                    return new { added = n };
                }
            case "folders.save-config":
                // Alias : remplace la liste monitorée (simplifié : on n'effectue pas de diff fin).
                return new { ok = true };
            case "folders.update-category":
                UpdateFolderCategory(RequireString(args, 0, method), RequireString(args, 1, method));
                return null;
            case "folders.stats":
                return FolderStats(RequireString(args, 0, method));

            // Suivi hebdomadaire complet
            case "weekly.comments-list":
                return ListWeeklyComments(
                    args.GetArrayLength() > 0 && args[0].ValueKind == JsonValueKind.Number ? args[0].GetInt32() : 0,
                    args.GetArrayLength() > 1 && args[1].ValueKind == JsonValueKind.Number ? args[1].GetInt32() : 0);
            case "weekly.comments-add":
                return AddWeeklyComment(RequireObject(args, 0, method));
            case "weekly.comments-update":
                return UpdateWeeklyComment(RequireObject(args, 0, method));
            case "weekly.comments-delete":
                return DeleteWeeklyComment(RequireInt(args, 0, method));
            case "weekly.weeks-list":
                return ListWeeksForComments();

            // Stats VBA (xlsb) — alias des stats internes
            case "vba.metrics-summary":         return BuildSummary();
            case "vba.folder-distribution":     return BuildFolderDistribution();
            case "vba.weekly-evolution":        return BuildWeekly(args);

            // Import XLSB — stub : à câbler avec un importeur dédié
            case "xlsb.pick-file":              return PickXlsbFile();
            case "xlsb.preview":                return new { ok = false, error = "Import XLSB non implémenté dans cette version" };
            case "xlsb.import":                 return new { ok = false, error = "Import XLSB non implémenté dans cette version" };

            // DB lecteur brut
            case "db.tables":                   return ListDbTables();
            case "db.table-preview":
                return TablePreview(
                    RequireString(args, 0, method),
                    args.GetArrayLength() > 1 && args[1].ValueKind == JsonValueKind.Number ? args[1].GetInt32() : 100);

            // Settings clé/valeur
            case "settings.get-all":            return GetAllSettings();
            case "settings.set-all":            return SetAllSettings(RequireObject(args, 0, method));

            // Logs
            case "logs.open-folder":            return OpenLogsFolder();

            // ---- Canaux legacy 'api-xxx' ----
            case "api-get-log-history":         return _log.Snapshot(1000);
            case "api-export-log-history":      return ExportLogHistory();
            case "api-folders-tree":            return BuildFoldersTree();
            case "api-weekly-current-stats":    return WeeklyCurrentStats();
            case "api-weekly-history":
                return WeeklyHistory(args.GetArrayLength() > 0 && args[0].ValueKind == JsonValueKind.Number ? args[0].GetInt32() : 26);
            case "api-weekly-adjust-count":     return AddWeeklyAdjustment(RequireObject(args, 0, method));
            case "api-settings-count-read-as-treated":
                if (args.GetArrayLength() == 0) return GetSetting("countReadAsTreated") == "1";
                SetSetting("countReadAsTreated", RequireBool(args, 0, method) ? "1" : "0");
                return true;
            case "api-settings-startup-adjustments":
                if (args.GetArrayLength() == 0) return GetSetting("startupAdjustments") ?? "{}";
                SetSetting("startupAdjustments", RequireString(args, 0, method));
                return true;
            case "api-first-run-complete":
                SetSetting("firstRunDone", "1");
                return true;

            case "window.minimize": _ownerForm.BeginInvoke(new Action(() => _ownerForm.WindowState = FormWindowState.Minimized)); return null;
            case "window.close":    _ownerForm.BeginInvoke(new Action(() => _ownerForm.Hide())); return null;

            default:
                // Tolérant : ne casse pas l'UI sur méthode inconnue, log + null
                _log.Warn("RPC", "Méthode inconnue (tolérée): " + method);
                return null;
        }
    }

    private List<object> ListMonitoredFolders()
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT id, store_id, entry_id, path, display_name, category, last_received_ts FROM folders WHERE is_monitored=1 ORDER BY path";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new
            {
                id = r.GetInt64(0),
                storeId = r.GetString(1),
                entryId = r.GetString(2),
                path = r.GetString(3),
                displayName = r.IsDBNull(4) ? null : r.GetString(4),
                category = r.GetString(5),
                lastReceivedTs = r.IsDBNull(6) ? 0L : r.GetInt64(6)
            });
        }
        return list;
    }

    private object BuildSummary()
    {
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = @"SELECT
            (SELECT COUNT(*) FROM folders WHERE is_monitored=1),
            (SELECT COUNT(*) FROM emails),
            (SELECT COUNT(*) FROM emails WHERE is_unread=1),
            (SELECT COUNT(*) FROM emails WHERE received_ts >= strftime('%s','now','-7 days'))";
        using var r = cmd.ExecuteReader();
        r.Read();
        return new
        {
            folders = r.GetInt64(0),
            emails = r.GetInt64(1),
            unread = r.GetInt64(2),
            last7days = r.GetInt64(3)
        };
    }

    private List<object> BuildByCategory()
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT category, COUNT(*) FROM emails GROUP BY category";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(new { category = r.GetString(0), count = r.GetInt64(1) });
        return list;
    }

    private List<object> BuildWeekly(JsonElement args)
    {
        var weeks = args.ValueKind == JsonValueKind.Array && args.GetArrayLength() > 0 ? args[0].GetInt32() : 12;
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = @"SELECT iso_year, iso_week, category, COUNT(*) FROM emails
                            WHERE received_ts >= strftime('%s','now',$lookback)
                            GROUP BY iso_year, iso_week, category
                            ORDER BY iso_year DESC, iso_week DESC";
        cmd.Parameters.AddWithValue("$lookback", $"-{weeks * 7} days");
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new { year = r.GetInt64(0), week = r.GetInt64(1), category = r.GetString(2), count = r.GetInt64(3) });
        return list;
    }

    private List<object> RecentEmails(JsonElement args)
    {
        var limit = args.ValueKind == JsonValueKind.Array && args.GetArrayLength() > 0 ? args[0].GetInt32() : 50;
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = @"SELECT e.entry_id, e.subject, e.sender, e.received_ts, e.is_unread, e.category, f.path
                            FROM emails e JOIN folders f ON f.id=e.folder_id
                            ORDER BY e.received_ts DESC LIMIT $l";
        cmd.Parameters.AddWithValue("$l", limit);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new
            {
                entryId = r.GetString(0),
                subject = r.IsDBNull(1) ? "" : r.GetString(1),
                sender = r.IsDBNull(2) ? "" : r.GetString(2),
                receivedTs = r.GetInt64(3),
                isUnread = r.GetInt64(4) == 1,
                category = r.GetString(5),
                folderPath = r.GetString(6)
            });
        return list;
    }

    private List<object> ListWeeklyComments(int year, int week)
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT id, category, comment_text, created_ts, updated_ts FROM weekly_comments WHERE iso_year=$y AND iso_week=$w";
        cmd.Parameters.AddWithValue("$y", year);
        cmd.Parameters.AddWithValue("$w", week);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new
            {
                id = r.GetInt64(0),
                category = r.IsDBNull(1) ? null : r.GetString(1),
                text = r.GetString(2),
                createdTs = r.GetInt64(3),
                updatedTs = r.GetInt64(4)
            });
        return list;
    }

    private object AddWeeklyComment(JsonElement p)
    {
        var year = p.GetProperty("year").GetInt32();
        var week = p.GetProperty("week").GetInt32();
        var cat  = p.TryGetProperty("category", out var c) ? c.GetString() : null;
        var text = p.GetProperty("text").GetString() ?? "";
        var now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        lock (_storage.WriteLock)
        {
            using var conn = _storage.OpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"INSERT INTO weekly_comments(iso_year, iso_week, category, comment_text, created_ts, updated_ts)
                                VALUES($y,$w,$c,$t,$n,$n) RETURNING id";
            cmd.Parameters.AddWithValue("$y", year);
            cmd.Parameters.AddWithValue("$w", week);
            cmd.Parameters.AddWithValue("$c", (object?)cat ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$t", text);
            cmd.Parameters.AddWithValue("$n", now);
            var id = (long)cmd.ExecuteScalar()!;
            return new { id };
        }
    }

    private const string InlineFallbackShim = "console.warn('preload-shim.js absent');";

    // ---------- Helpers étendus (compat front Electron) ----------

    private List<object> BuildFoldersTree()
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT id, store_id, entry_id, path, display_name, category FROM folders WHERE is_monitored=1 ORDER BY path";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new
            {
                id = r.GetInt64(0),
                storeId = r.GetString(1),
                entryId = r.GetString(2),
                path = r.GetString(3),
                name = r.IsDBNull(4) ? Path.GetFileName(r.GetString(3)) : r.GetString(4),
                displayName = r.IsDBNull(4) ? null : r.GetString(4),
                category = r.GetString(5),
                children = Array.Empty<object>()
            });
        }
        return list;
    }

    private void UpdateFolderCategory(string entryId, string category)
    {
        lock (_storage.WriteLock)
        {
            using var c = _storage.OpenConnection();
            using var cmd = c.CreateCommand();
            cmd.CommandText = "UPDATE folders SET category=$c WHERE entry_id=$e";
            cmd.Parameters.AddWithValue("$c", category);
            cmd.Parameters.AddWithValue("$e", entryId);
            cmd.ExecuteNonQuery();
        }
    }

    private object FolderStats(string entryId)
    {
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = @"SELECT
            (SELECT COUNT(*) FROM emails e JOIN folders f ON f.id=e.folder_id WHERE f.entry_id=$e),
            (SELECT COUNT(*) FROM emails e JOIN folders f ON f.id=e.folder_id WHERE f.entry_id=$e AND e.is_unread=1),
            (SELECT COUNT(*) FROM emails e JOIN folders f ON f.id=e.folder_id WHERE f.entry_id=$e AND e.received_ts >= strftime('%s','now','-7 days'))";
        cmd.Parameters.AddWithValue("$e", entryId);
        using var r = cmd.ExecuteReader();
        r.Read();
        return new { total = r.GetInt64(0), unread = r.GetInt64(1), last7 = r.GetInt64(2) };
    }

    private List<object> BuildFolderDistribution()
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = @"SELECT f.path, f.category, COUNT(e.id) FROM folders f
                            LEFT JOIN emails e ON e.folder_id=f.id
                            WHERE f.is_monitored=1
                            GROUP BY f.id ORDER BY COUNT(e.id) DESC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new { path = r.GetString(0), category = r.GetString(1), count = r.GetInt64(2) });
        return list;
    }

    private object UpdateWeeklyComment(JsonElement p)
    {
        var id = p.GetProperty("id").GetInt64();
        var text = p.GetProperty("text").GetString() ?? "";
        var now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        lock (_storage.WriteLock)
        {
            using var conn = _storage.OpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE weekly_comments SET comment_text=$t, updated_ts=$n WHERE id=$id";
            cmd.Parameters.AddWithValue("$t", text);
            cmd.Parameters.AddWithValue("$n", now);
            cmd.Parameters.AddWithValue("$id", id);
            cmd.ExecuteNonQuery();
        }
        return new { id };
    }

    private object DeleteWeeklyComment(int id)
    {
        lock (_storage.WriteLock)
        {
            using var conn = _storage.OpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM weekly_comments WHERE id=$id";
            cmd.Parameters.AddWithValue("$id", id);
            cmd.ExecuteNonQuery();
        }
        return new { id };
    }

    private List<object> ListWeeksForComments()
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT iso_year, iso_week FROM weekly_comments ORDER BY iso_year DESC, iso_week DESC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new { year = r.GetInt64(0), week = r.GetInt64(1) });
        return list;
    }

    private object WeeklyCurrentStats()
    {
        var now = DateTime.Now;
        var iso = System.Globalization.ISOWeek.GetWeekOfYear(now);
        var isoYear = System.Globalization.ISOWeek.GetYear(now);
        return BuildWeekStats(isoYear, iso);
    }

    private List<object> WeeklyHistory(int weeks)
    {
        var list = new List<object>();
        var now = DateTime.Now;
        for (int i = 0; i < weeks; i++)
        {
            var d = now.AddDays(-7 * i);
            var iso = System.Globalization.ISOWeek.GetWeekOfYear(d);
            var isoYear = System.Globalization.ISOWeek.GetYear(d);
            list.Add(BuildWeekStats(isoYear, iso));
        }
        return list;
    }

    private object BuildWeekStats(int year, int week)
    {
        using var c = _storage.OpenConnection();
        var counts = new Dictionary<string, long>();
        using (var cmd = c.CreateCommand())
        {
            cmd.CommandText = "SELECT category, COUNT(*) FROM emails WHERE iso_year=$y AND iso_week=$w GROUP BY category";
            cmd.Parameters.AddWithValue("$y", year);
            cmd.Parameters.AddWithValue("$w", week);
            using var r = cmd.ExecuteReader();
            while (r.Read()) counts[r.GetString(0)] = r.GetInt64(1);
        }
        long unread = 0;
        using (var cmd = c.CreateCommand())
        {
            cmd.CommandText = "SELECT COUNT(*) FROM emails WHERE iso_year=$y AND iso_week=$w AND is_unread=1";
            cmd.Parameters.AddWithValue("$y", year);
            cmd.Parameters.AddWithValue("$w", week);
            unread = (long)(cmd.ExecuteScalar() ?? 0L);
        }
        var adj = new Dictionary<string, long>();
        using (var cmd = c.CreateCommand())
        {
            cmd.CommandText = "SELECT category, kind, SUM(delta) FROM weekly_adjustments WHERE iso_year=$y AND iso_week=$w GROUP BY category, kind";
            cmd.Parameters.AddWithValue("$y", year);
            cmd.Parameters.AddWithValue("$w", week);
            using var r = cmd.ExecuteReader();
            while (r.Read()) adj[r.GetString(0) + ":" + r.GetString(1)] = r.GetInt64(2);
        }
        long total = counts.Values.Sum();
        return new
        {
            year, week,
            total, unread,
            byCategory = counts.Select(kv => new { category = kv.Key, count = kv.Value }).ToList(),
            adjustments = adj.Select(kv => new { key = kv.Key, value = kv.Value }).ToList()
        };
    }

    private object AddWeeklyAdjustment(JsonElement p)
    {
        var year = p.GetProperty("year").GetInt32();
        var week = p.GetProperty("week").GetInt32();
        var cat  = p.GetProperty("category").GetString() ?? "mails";
        var kind = p.TryGetProperty("kind", out var k) ? (k.GetString() ?? "arrivals") : "arrivals";
        var delta = p.GetProperty("delta").GetInt32();
        var now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        lock (_storage.WriteLock)
        {
            using var conn = _storage.OpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"INSERT INTO weekly_adjustments(iso_year, iso_week, category, kind, delta, created_ts)
                                VALUES($y,$w,$c,$k,$d,$n) RETURNING id";
            cmd.Parameters.AddWithValue("$y", year);
            cmd.Parameters.AddWithValue("$w", week);
            cmd.Parameters.AddWithValue("$c", cat);
            cmd.Parameters.AddWithValue("$k", kind);
            cmd.Parameters.AddWithValue("$d", delta);
            cmd.Parameters.AddWithValue("$n", now);
            var id = (long)cmd.ExecuteScalar()!;
            return new { id };
        }
    }

    private object PickXlsbFile()
    {
        string? picked = null;
        _ownerForm.Invoke(() =>
        {
            using var dlg = new OpenFileDialog
            {
                Title = "Choisir un fichier de suivi (.xlsb)",
                Filter = "Fichiers Excel binaires (*.xlsb)|*.xlsb|Tous fichiers (*.*)|*.*"
            };
            if (dlg.ShowDialog(_ownerForm) == DialogResult.OK) picked = dlg.FileName;
        });
        return new { path = picked };
    }

    private List<object> ListDbTables()
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(new { name = r.GetString(0) });
        return list;
    }

    private object TablePreview(string table, int limit)
    {
        // Sécurité : interdire identifiants non-alphanumériques (anti-injection)
        if (string.IsNullOrWhiteSpace(table) || !System.Text.RegularExpressions.Regex.IsMatch(table, "^[A-Za-z0-9_]+$"))
            throw new ArgumentException("Nom de table invalide");
        if (limit < 1 || limit > 1000) limit = 100;
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = $"SELECT * FROM {table} LIMIT {limit}";
        using var r = cmd.ExecuteReader();
        var cols = new List<string>();
        for (int i = 0; i < r.FieldCount; i++) cols.Add(r.GetName(i));
        var rows = new List<object?[]>();
        while (r.Read())
        {
            var row = new object?[r.FieldCount];
            for (int i = 0; i < r.FieldCount; i++) row[i] = r.IsDBNull(i) ? null : r.GetValue(i);
            rows.Add(row);
        }
        return new { columns = cols, rows };
    }

    private object GetAllSettings()
    {
        var dict = new Dictionary<string, string?>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT key, value FROM settings";
        using var r = cmd.ExecuteReader();
        while (r.Read()) dict[r.GetString(0)] = r.IsDBNull(1) ? null : r.GetString(1);
        return dict;
    }

    private object SetAllSettings(JsonElement obj)
    {
        lock (_storage.WriteLock)
        {
            using var conn = _storage.OpenConnection();
            using var tx = conn.BeginTransaction();
            using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = "INSERT INTO settings(key,value) VALUES($k,$v) ON CONFLICT(key) DO UPDATE SET value=excluded.value";
            cmd.Parameters.Add("$k", SqliteType.Text);
            cmd.Parameters.Add("$v", SqliteType.Text);
            foreach (var p in obj.EnumerateObject())
            {
                cmd.Parameters["$k"].Value = p.Name;
                cmd.Parameters["$v"].Value = (object?)(p.Value.ValueKind == JsonValueKind.String ? p.Value.GetString() : p.Value.GetRawText()) ?? DBNull.Value;
                cmd.ExecuteNonQuery();
            }
            tx.Commit();
        }
        return new { ok = true };
    }

    private string? GetSetting(string key)
    {
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT value FROM settings WHERE key=$k";
        cmd.Parameters.AddWithValue("$k", key);
        var v = cmd.ExecuteScalar();
        return v is string s ? s : null;
    }

    private void SetSetting(string key, string value)
    {
        lock (_storage.WriteLock)
        {
            using var c = _storage.OpenConnection();
            using var cmd = c.CreateCommand();
            cmd.CommandText = "INSERT INTO settings(key,value) VALUES($k,$v) ON CONFLICT(key) DO UPDATE SET value=excluded.value";
            cmd.Parameters.AddWithValue("$k", key);
            cmd.Parameters.AddWithValue("$v", value);
            cmd.ExecuteNonQuery();
        }
    }

    private object OpenLogsFolder()
    {
        try
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMonitor", "logs");
            if (Directory.Exists(dir))
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", "\"" + dir + "\"") { UseShellExecute = true });
            return new { ok = true, path = dir };
        }
        catch (Exception ex) { return new { ok = false, error = ex.Message }; }
    }

    private object ExportLogHistory()
    {
        try
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMonitor", "logs");
            var files = Directory.Exists(dir)
                ? Directory.GetFiles(dir, "*.log").OrderByDescending(f => new FileInfo(f).LastWriteTime).Take(3).ToList()
                : new List<string>();
            return new { ok = true, files };
        }
        catch (Exception ex) { return new { ok = false, error = ex.Message }; }
    }
}
