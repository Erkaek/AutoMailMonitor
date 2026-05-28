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

    private readonly System.Collections.Concurrent.ConcurrentQueue<string> _eventQueue = new();
    private System.Threading.Timer? _statusTimer;

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

        _monitor.OnMailEvent += e =>
        {
            var t = e.Type.ToString().ToLowerInvariant();
            if (t == "add")    { PostEvent("email.new", e); PostEvent("email.realtime-new", e); }
            else               { PostEvent("email.update", e); PostEvent("email.realtime-update", e); }
        };
        _monitor.OnStatsChanged += () =>
        {
            PostEvent("stats.update", null);
            PostEvent("stats.cache-invalidated", null);
            PostEvent("weekly.stats-updated", null);
            PostEvent("monitoring.status", BuildMonitoringStatus());
        };
        _log.OnEntry += entry => PostEvent("logs.entry", ToLegacyLog(entry));
        _updater.OnAvailable += v => PostEvent("update.available", new { version = v });
        _updater.OnProgress += p => PostEvent("update.progress", new { percent = p });
        _updater.OnReadyToInstall += () => PostEvent("update.ready", null);

        // Broadcast initial + périodique du statut (sidebar "Initialisation..." → "Connecté")
        _statusTimer = new System.Threading.Timer(_ =>
        {
            try
            {
                var connected = _monitor.Outlook.IsConnected;
                PostEvent("monitoring.status", BuildMonitoringStatus());
                PostEvent("outlook.status", new { status = connected, connected, error = (string?)null });
                if (connected) PostEvent("com.listening-started", new { ok = true });
            }
            catch { }
        }, null, TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(5));
    }

    private object BuildMonitoringStatus()
    {
        long folders = 0;
        try
        {
            using var c = _storage.OpenConnection();
            using var cmd = c.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM folders WHERE is_monitored=1";
            folders = (long)(cmd.ExecuteScalar() ?? 0L);
        }
        catch { }
        return new
        {
            running = _monitor.IsRunning,
            active = _monitor.IsRunning,
            foldersMonitored = folders,
            connected = _monitor.Outlook.IsConnected
        };
    }

    private void PostEvent(string name, object? data)
    {
        try
        {
            var json = JsonSerializer.Serialize(new { @event = name, data }, Json);
            if (_web.CoreWebView2 is null || !_ownerForm.IsHandleCreated)
            {
                _eventQueue.Enqueue(json);
                return;
            }
            // Drain queue d'abord
            while (_eventQueue.TryDequeue(out var pending))
            {
                var snap = pending;
                _ownerForm.BeginInvoke(new Action(() => { try { _web.CoreWebView2?.PostWebMessageAsJson(snap); } catch { } }));
            }
            _ownerForm.BeginInvoke(new Action(() => { try { _web.CoreWebView2?.PostWebMessageAsJson(json); } catch { } }));
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

            case "monitoring.status": return BuildMonitoringStatus();
            case "outlook.status": return new { status = _monitor.Outlook.IsConnected, connected = _monitor.Outlook.IsConnected, error = (string?)null };

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

            case "folders.list-monitored": return ListMonitoredFoldersForConfig();
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
            case "stats.by-category": return new { categories = BuildByCategoryMap() };

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

            case "outlook.folder-tree-from-path":
                {
                    var rootPath = RequireString(args, 0, method);
                    int maxDepth = 4;
                    if (args.GetArrayLength() > 1 && args[1].ValueKind == JsonValueKind.Number)
                        maxDepth = Math.Max(1, Math.Min(12, args[1].GetInt32()));
                    var tree = await _monitor.Outlook.GetFolderTreeFromPathAsync(rootPath, maxDepth);
                    if (!tree.Success || tree.Root is null)
                        return new { success = false, error = tree.Error ?? "Échec de la résolution du chemin" };
                    return new
                    {
                        success = true,
                        root = tree.Root,
                        store = new { id = tree.StoreId, name = tree.StoreName, smtp = tree.StoreSmtp }
                    };
                }

            // Folders config bulk + update
            case "folders.list-stats-shape":
                return new { stats = BuildFolderStatsList() };
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
                // Compat: peut être appelé avec un entryId (string) ou un objet d'options ({force:true})
                if (args.ValueKind == JsonValueKind.Array && args.GetArrayLength() > 0 && args[0].ValueKind == JsonValueKind.String)
                    return FolderStats(args[0].GetString() ?? "");
                return new { stats = BuildFolderStatsList() };

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
                return TablePreviewFromArgs(args);

            // Settings clé/valeur
            case "settings.get-all":            return GetAllSettings();
            case "settings.set-all":            return SetAllSettings(RequireObject(args, 0, method));

            // Logs
            case "logs.open-folder":            return OpenLogsFolder();

            // ---- Canaux legacy 'api-xxx' ----
            case "api-get-log-history":         return FilterLogHistory(args);
            case "api-export-log-history":      return ExportLogHistory(args);
            case "api-folders-tree":            return BuildFoldersTree();
            case "api-weekly-current-stats":    return WeeklyCurrentStats();
            case "api-weekly-history":          return WeeklyHistoryPaged(args);
            case "api-weekly-adjust-count":     return AddWeeklyAdjustment(RequireObject(args, 0, method));
            case "api-settings-count-read-as-treated":
                return SettingsCountReadAsTreated(args);
            case "api-settings-startup-adjustments":
                if (args.GetArrayLength() == 0) return new { success = true, value = GetSetting("startupAdjustments") ?? "{}" };
                SetSetting("startupAdjustments", args[0].ValueKind == JsonValueKind.String ? (args[0].GetString() ?? "") : args[0].GetRawText());
                return new { success = true };
            case "api-first-run-complete":
                SetSetting("firstRunDone", "1");
                return new { success = true };
            case "api-force-full-resync":
                return new { success = true };

            case "window.minimize": _ownerForm.BeginInvoke(new Action(() => _ownerForm.WindowState = FormWindowState.Minimized)); return null;
            case "window.close":    _ownerForm.BeginInvoke(new Action(() => _ownerForm.Hide())); return null;

            default:
                // Tolérant : ne casse pas l'UI sur méthode inconnue, log + null
                _log.Warn("RPC", "Méthode inconnue (tolérée): " + method);
                return null;
        }
    }

    private object ListMonitoredFoldersForConfig()
    {
        var folderCategories = new Dictionary<string, string>();
        var folders = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT id, store_id, entry_id, path, display_name, category, last_received_ts FROM folders WHERE is_monitored=1 ORDER BY path";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var path = r.GetString(3);
            var category = r.GetString(5);
            folderCategories[path] = category;
            folders.Add(new
            {
                id = r.GetInt64(0),
                storeId = r.GetString(1),
                entryId = r.GetString(2),
                path,
                displayName = r.IsDBNull(4) ? null : r.GetString(4),
                category,
                lastReceivedTs = r.IsDBNull(6) ? 0L : r.GetInt64(6)
            });
        }
        return new { success = true, folderCategories, folders };
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
            (SELECT COUNT(*) FROM emails WHERE received_ts >= strftime('%s','now','-7 days')),
            (SELECT COUNT(*) FROM emails WHERE date(received_ts,'unixepoch','localtime') = date('now','localtime'))";
        using var r = cmd.ExecuteReader();
        r.Read();
        var folders = r.GetInt64(0);
        var emails = r.GetInt64(1);
        var unread = r.GetInt64(2);
        var last7 = r.GetInt64(3);
        var today = r.GetInt64(4);
        return new
        {
            folders,
            totalEmails = emails,
            unreadTotal = unread,
            emailsToday = today,
            last7days = last7,
            // alias historiques
            emails,
            unread
        };
    }

    private Dictionary<string, long> BuildByCategoryMap()
    {
        var map = new Dictionary<string, long>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT COALESCE(category,'mails'), COUNT(*) FROM emails GROUP BY category";
        using var r = cmd.ExecuteReader();
        while (r.Read()) map[r.GetString(0)] = r.GetInt64(1);
        return map;
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
        int limit = 200;
        if (args.ValueKind == JsonValueKind.Array && args.GetArrayLength() > 0)
        {
            var a0 = args[0];
            if (a0.ValueKind == JsonValueKind.Number) limit = a0.GetInt32();
            else if (a0.ValueKind == JsonValueKind.Object && a0.TryGetProperty("limit", out var lim) && lim.ValueKind == JsonValueKind.Number)
                limit = lim.GetInt32();
        }
        if (limit < 1 || limit > 2000) limit = 200;
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = @"SELECT e.entry_id, e.subject, e.sender, e.received_ts, e.is_unread, e.category, f.path, f.display_name
                            FROM emails e JOIN folders f ON f.id=e.folder_id
                            ORDER BY e.received_ts DESC LIMIT $l";
        cmd.Parameters.AddWithValue("$l", limit);
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var entryId = r.GetString(0);
            var subject = r.IsDBNull(1) ? "" : r.GetString(1);
            var sender = r.IsDBNull(2) ? "" : r.GetString(2);
            var ts = r.GetInt64(3);
            var isUnread = r.GetInt64(4) == 1;
            var category = r.IsDBNull(5) ? "" : r.GetString(5);
            var path = r.GetString(6);
            var name = r.IsDBNull(7) ? Path.GetFileName(path) : r.GetString(7);
            var iso = DateTimeOffset.FromUnixTimeSeconds(ts).ToLocalTime().ToString("yyyy-MM-ddTHH:mm:sszzz");
            list.Add(new
            {
                entry_id = entryId,
                entryId,
                subject,
                sender,
                sender_name = sender,
                sender_email = sender,
                received_ts = ts,
                receivedTs = ts,
                received_time = iso,
                ReceivedTime = iso,
                is_unread = isUnread,
                isUnread,
                is_read = !isUnread,
                UnRead = isUnread,
                category,
                folder_path = path,
                folderPath = path,
                folder_name = name,
                Folder = name,
                FolderPath = path
            });
        }
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

    private object BuildFoldersTree()
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = @"SELECT f.id, f.store_id, f.entry_id, f.path, f.display_name, f.category,
                                   (SELECT COUNT(*) FROM emails e WHERE e.folder_id=f.id) AS total,
                                   (SELECT COUNT(*) FROM emails e WHERE e.folder_id=f.id AND e.is_unread=1) AS unread
                            FROM folders f WHERE f.is_monitored=1 ORDER BY f.path";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var path = r.GetString(3);
            var name = r.IsDBNull(4) ? Path.GetFileName(path) : r.GetString(4);
            list.Add(new
            {
                id = r.GetInt64(0),
                storeId = r.GetString(1),
                entryId = r.GetString(2),
                path,
                name,
                displayName = name,
                folder_name = name,
                folder_path = path,
                category = r.GetString(5),
                total = r.GetInt64(6),
                unread = r.GetInt64(7),
                children = Array.Empty<object>()
            });
        }
        return new { folders = list, success = true };
    }

    private List<object> BuildFolderStatsList()
    {
        var list = new List<object>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = @"SELECT f.path, f.display_name, f.category,
                                   (SELECT COUNT(*) FROM emails e WHERE e.folder_id=f.id) AS total,
                                   (SELECT COUNT(*) FROM emails e WHERE e.folder_id=f.id AND e.is_unread=1) AS unread
                            FROM folders f WHERE f.is_monitored=1 ORDER BY f.path";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var path = r.GetString(0);
            var name = r.IsDBNull(1) ? Path.GetFileName(path) : r.GetString(1);
            list.Add(new
            {
                path,
                folder_name = name,
                folder_path = path,
                category = r.GetString(2),
                total = r.GetInt64(3),
                unread = r.GetInt64(4)
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

    private static readonly string[] WeeklyCategoryDisplay = { "Déclarations", "Règlements", "Mails simples" };
    private static readonly string[] WeeklyCategoryKeys = { "declarations", "reglements", "mails_simples" };

    private static (DateTime Monday, DateTime Sunday) WeekRange(int year, int week)
    {
        var monday = System.Globalization.ISOWeek.ToDateTime(year, week, DayOfWeek.Monday);
        return (monday, monday.AddDays(6));
    }

    private object WeeklyCurrentStats()
    {
        var now = DateTime.Now;
        var iso = System.Globalization.ISOWeek.GetWeekOfYear(now);
        var isoYear = System.Globalization.ISOWeek.GetYear(now);
        return BuildCurrentWeekResponse(isoYear, iso);
    }

    private object BuildCurrentWeekResponse(int year, int week)
    {
        var stats = BuildPerCategoryStats(year, week);
        var (mon, sun) = WeekRange(year, week);
        var categories = new Dictionary<string, object>();
        for (int i = 0; i < WeeklyCategoryKeys.Length; i++)
        {
            var s = stats[WeeklyCategoryKeys[i]];
            categories[WeeklyCategoryDisplay[i]] = new
            {
                received = s.received,
                treated = s.treated,
                adjustments = s.adjustments,
                total = s.total
            };
        }
        return new
        {
            success = true,
            weekInfo = new
            {
                displayName = $"S{week:00} {year}",
                startDate = mon.ToString("dd/MM/yyyy"),
                endDate = sun.ToString("dd/MM/yyyy"),
                identifier = $"S{week:00}-{year}",
                isoYear = year,
                isoWeek = week
            },
            categories
        };
    }

    private Dictionary<string, (long received, long treated, long adjustments, long total)> BuildPerCategoryStats(int year, int week)
    {
        var result = new Dictionary<string, (long received, long treated, long adjustments, long total)>();
        var countReadAsTreated = GetSetting("countReadAsTreated") == "1";
        using var c = _storage.OpenConnection();
        for (int i = 0; i < WeeklyCategoryKeys.Length; i++)
        {
            var key = WeeklyCategoryKeys[i];
            long received = 0, treated = 0, adjustments = 0;
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(*) FROM emails WHERE iso_year=$y AND iso_week=$w AND LOWER(COALESCE(category,''))=$c";
                cmd.Parameters.AddWithValue("$y", year);
                cmd.Parameters.AddWithValue("$w", week);
                cmd.Parameters.AddWithValue("$c", key);
                received = (long)(cmd.ExecuteScalar() ?? 0L);
            }
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = countReadAsTreated
                    ? "SELECT COUNT(*) FROM emails WHERE iso_year=$y AND iso_week=$w AND LOWER(COALESCE(category,''))=$c AND is_unread=0"
                    : "SELECT 0";
                cmd.Parameters.AddWithValue("$y", year);
                cmd.Parameters.AddWithValue("$w", week);
                cmd.Parameters.AddWithValue("$c", key);
                treated = (long)(cmd.ExecuteScalar() ?? 0L);
            }
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "SELECT COALESCE(SUM(delta),0) FROM weekly_adjustments WHERE iso_year=$y AND iso_week=$w AND LOWER(category)=$c AND kind='treated'";
                cmd.Parameters.AddWithValue("$y", year);
                cmd.Parameters.AddWithValue("$w", week);
                cmd.Parameters.AddWithValue("$c", key);
                adjustments = (long)(cmd.ExecuteScalar() ?? 0L);
            }
            result[key] = (received, treated, adjustments, received);
        }
        return result;
    }

    private object WeeklyHistoryPaged(JsonElement args)
    {
        int page = 1, pageSize = 12;
        if (args.ValueKind == JsonValueKind.Array && args.GetArrayLength() > 0 && args[0].ValueKind == JsonValueKind.Object)
        {
            var a = args[0];
            if (a.TryGetProperty("page", out var pe) && pe.ValueKind == JsonValueKind.Number) page = Math.Max(1, pe.GetInt32());
            if (a.TryGetProperty("pageSize", out var ps) && ps.ValueKind == JsonValueKind.Number) pageSize = Math.Max(1, ps.GetInt32());
            else if (a.TryGetProperty("limit", out var lm) && lm.ValueKind == JsonValueKind.Number) pageSize = Math.Max(1, lm.GetInt32());
        }
        // Liste toutes les semaines distinctes pr\u00e9sentes en BDD
        var weeks = new List<(int year, int week)>();
        using (var c = _storage.OpenConnection())
        using (var cmd = c.CreateCommand())
        {
            cmd.CommandText = @"SELECT iso_year, iso_week FROM (
                                  SELECT iso_year, iso_week FROM emails
                                  UNION SELECT iso_year, iso_week FROM weekly_adjustments
                                ) GROUP BY iso_year, iso_week ORDER BY iso_year DESC, iso_week DESC";
            using var r = cmd.ExecuteReader();
            while (r.Read()) weeks.Add((r.GetInt32(0), r.GetInt32(1)));
        }
        if (weeks.Count == 0)
        {
            // Au moins la semaine courante
            var now = DateTime.Now;
            weeks.Add((System.Globalization.ISOWeek.GetYear(now), System.Globalization.ISOWeek.GetWeekOfYear(now)));
        }
        var totalWeeks = weeks.Count;
        var totalPages = Math.Max(1, (int)Math.Ceiling(totalWeeks / (double)pageSize));
        if (page > totalPages) page = totalPages;
        var slice = weeks.Skip((page - 1) * pageSize).Take(pageSize);

        var data = new List<object>();
        foreach (var (year, week) in slice)
        {
            var stats = BuildPerCategoryStats(year, week);
            var (mon, sun) = WeekRange(year, week);
            var cats = new List<object>();
            for (int i = 0; i < WeeklyCategoryKeys.Length; i++)
            {
                var s = stats[WeeklyCategoryKeys[i]];
                cats.Add(new
                {
                    name = WeeklyCategoryDisplay[i],
                    received = s.received,
                    treated = s.treated,
                    adjustments = s.adjustments,
                    stockEndWeek = Math.Max(0, s.received - s.treated - s.adjustments)
                });
            }
            data.Add(new
            {
                week_year = year,
                week_number = week,
                week_identifier = $"S{week:00}-{year}",
                weekDisplay = $"S{week:00} - {year}",
                dateRange = $"{mon:dd/MM/yyyy} - {sun:dd/MM/yyyy}",
                evolution = new { trend = "stable", percent = 0 },
                categories = cats
            });
        }
        return new
        {
            success = true,
            data,
            page,
            pageSize,
            totalWeeks,
            totalPages
        };
    }

    private object AddWeeklyAdjustment(JsonElement p)
    {
        // Nouveau contrat: { weekIdentifier, folderType, adjustmentValue, adjustmentType }
        int year, week;
        string cat;
        string kind = "treated";
        int delta;
        if (p.TryGetProperty("weekIdentifier", out var wid) && wid.ValueKind == JsonValueKind.String)
        {
            var m = System.Text.RegularExpressions.Regex.Match(wid.GetString() ?? "", @"S\s*(\d{1,2}).*?(\d{4})");
            if (!m.Success) return new { success = false, error = "weekIdentifier invalide" };
            week = int.Parse(m.Groups[1].Value);
            year = int.Parse(m.Groups[2].Value);
            cat = p.TryGetProperty("folderType", out var ft) ? (ft.GetString() ?? "mails_simples") : "mails_simples";
            delta = p.TryGetProperty("adjustmentValue", out var av) && av.ValueKind == JsonValueKind.Number ? av.GetInt32() : 0;
            if (p.TryGetProperty("adjustmentType", out var at) && at.ValueKind == JsonValueKind.String)
            {
                var atv = at.GetString() ?? "";
                kind = atv == "manual_adjustments" ? "treated" : atv;
            }
        }
        else
        {
            // Ancien contrat
            year = p.GetProperty("year").GetInt32();
            week = p.GetProperty("week").GetInt32();
            cat  = p.TryGetProperty("category", out var ca) ? (ca.GetString() ?? "mails_simples") : "mails_simples";
            kind = p.TryGetProperty("kind", out var k) ? (k.GetString() ?? "treated") : "treated";
            delta = p.GetProperty("delta").GetInt32();
        }
        if (delta == 0) return new { success = false, error = "Valeur d'ajustement nulle" };
        var now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        long id;
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
            id = (long)cmd.ExecuteScalar()!;
        }
        return new { success = true, id };
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

    private object ListDbTables()
    {
        var list = new List<string>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return new { success = true, tables = list };
    }

    private object TablePreviewFromArgs(JsonElement args)
    {
        string? table = null;
        int limit = 100;
        if (args.ValueKind == JsonValueKind.Array && args.GetArrayLength() > 0)
        {
            var a0 = args[0];
            if (a0.ValueKind == JsonValueKind.String) table = a0.GetString();
            else if (a0.ValueKind == JsonValueKind.Object)
            {
                if (a0.TryGetProperty("table", out var t) && t.ValueKind == JsonValueKind.String) table = t.GetString();
                if (a0.TryGetProperty("limit", out var l) && l.ValueKind == JsonValueKind.Number) limit = l.GetInt32();
            }
            if (args.GetArrayLength() > 1 && args[1].ValueKind == JsonValueKind.Number) limit = args[1].GetInt32();
        }
        if (string.IsNullOrWhiteSpace(table)) return new { success = false, error = "Table manquante" };
        try { return new { success = true, data = TablePreview(table!, limit) }; }
        catch (Exception ex) { return new { success = false, error = ex.Message }; }
    }

    private object TablePreview(string table, int limit)
    {
        if (string.IsNullOrWhiteSpace(table) || !System.Text.RegularExpressions.Regex.IsMatch(table, "^[A-Za-z0-9_]+$"))
            throw new ArgumentException("Nom de table invalide");
        if (limit < 1 || limit > 1000) limit = 100;
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = $"SELECT * FROM {table} LIMIT {limit}";
        using var r = cmd.ExecuteReader();
        var cols = new List<string>();
        for (int i = 0; i < r.FieldCount; i++) cols.Add(r.GetName(i));
        var rows = new List<Dictionary<string, object?>>();
        while (r.Read())
        {
            var row = new Dictionary<string, object?>(r.FieldCount);
            for (int i = 0; i < r.FieldCount; i++) row[cols[i]] = r.IsDBNull(i) ? null : r.GetValue(i);
            rows.Add(row);
        }
        return new { columns = cols, rows };
    }

    private object GetAllSettings()
    {
        var raw = new Dictionary<string, string?>();
        using var c = _storage.OpenConnection();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT key, value FROM settings";
        using var r = cmd.ExecuteReader();
        while (r.Read()) raw[r.GetString(0)] = r.IsDBNull(1) ? null : r.GetString(1);

        bool TryBool(string k, bool d = false) => raw.TryGetValue(k, out var v) && (v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase)) ? true : d;
        int TryInt(string k, int d) => raw.TryGetValue(k, out var v) && int.TryParse(v, out var n) ? n : d;
        string TryStr(string k, string d) => raw.TryGetValue(k, out var v) && !string.IsNullOrEmpty(v) ? v! : d;

        var settings = new
        {
            monitoring = new
            {
                treatReadEmailsAsProcessed = TryBool("countReadAsTreated"),
                scanInterval = TryInt("scanInterval", 30000),
                autoStart = TryBool("autoStart", true)
            },
            ui = new
            {
                language = TryStr("lang", "fr"),
                emailsLimit = TryInt("emailsLimit", 20),
                theme = TryStr("theme", "light"),
                tabs = new
                {
                    dashboard = TryBool("tab.dashboard", true),
                    emails = TryBool("tab.emails", true),
                    weekly = TryBool("tab.weekly", true),
                    personalPerformance = TryBool("tab.personalPerformance", true),
                    importActivity = TryBool("tab.importActivity", true),
                    monitoring = TryBool("tab.monitoring", true),
                    db = TryBool("tab.db", false)
                }
            },
            database = new
            {
                purgeOldDataAfterDays = TryInt("purgeOldDataAfterDays", 365),
                enableEventLogging = TryBool("enableEventLogging", true)
            },
            notifications = new
            {
                showStartupNotification = TryBool("showStartupNotification", true),
                showMonitoringStatus = TryBool("showMonitoringStatus", true),
                enableDesktopNotifications = TryBool("enableDesktopNotifications", false)
            }
        };
        return new { success = true, settings };
    }

    private object SetAllSettings(JsonElement obj)
    {
        // Flatten object → clés plates
        var flat = new Dictionary<string, string?>();
        void Walk(string prefix, JsonElement el)
        {
            switch (el.ValueKind)
            {
                case JsonValueKind.Object:
                    foreach (var p in el.EnumerateObject()) Walk(prefix.Length == 0 ? p.Name : prefix + "." + p.Name, p.Value);
                    break;
                case JsonValueKind.String: flat[prefix] = el.GetString(); break;
                case JsonValueKind.Number: flat[prefix] = el.GetRawText(); break;
                case JsonValueKind.True: flat[prefix] = "1"; break;
                case JsonValueKind.False: flat[prefix] = "0"; break;
                case JsonValueKind.Null: flat[prefix] = null; break;
                default: flat[prefix] = el.GetRawText(); break;
            }
        }
        Walk("", obj);
        // Mapping vers clés legacy
        var map = new Dictionary<string, string>
        {
            ["monitoring.treatReadEmailsAsProcessed"] = "countReadAsTreated",
            ["monitoring.scanInterval"] = "scanInterval",
            ["monitoring.autoStart"] = "autoStart",
            ["ui.language"] = "lang",
            ["ui.emailsLimit"] = "emailsLimit",
            ["ui.theme"] = "theme",
            ["ui.tabs.dashboard"] = "tab.dashboard",
            ["ui.tabs.emails"] = "tab.emails",
            ["ui.tabs.weekly"] = "tab.weekly",
            ["ui.tabs.personalPerformance"] = "tab.personalPerformance",
            ["ui.tabs.importActivity"] = "tab.importActivity",
            ["ui.tabs.monitoring"] = "tab.monitoring",
            ["ui.tabs.db"] = "tab.db",
            ["database.purgeOldDataAfterDays"] = "purgeOldDataAfterDays",
            ["database.enableEventLogging"] = "enableEventLogging",
            ["notifications.showStartupNotification"] = "showStartupNotification",
            ["notifications.showMonitoringStatus"] = "showMonitoringStatus",
            ["notifications.enableDesktopNotifications"] = "enableDesktopNotifications"
        };
        lock (_storage.WriteLock)
        {
            using var conn = _storage.OpenConnection();
            using var tx = conn.BeginTransaction();
            using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = "INSERT INTO settings(key,value) VALUES($k,$v) ON CONFLICT(key) DO UPDATE SET value=excluded.value";
            cmd.Parameters.Add("$k", SqliteType.Text);
            cmd.Parameters.Add("$v", SqliteType.Text);
            foreach (var kv in flat)
            {
                var dbKey = map.TryGetValue(kv.Key, out var mk) ? mk : kv.Key;
                cmd.Parameters["$k"].Value = dbKey;
                cmd.Parameters["$v"].Value = (object?)kv.Value ?? DBNull.Value;
                cmd.ExecuteNonQuery();
            }
            tx.Commit();
        }
        return new { success = true };
    }

    private object SettingsCountReadAsTreated(JsonElement args)
    {
        // Aucun arg → lecture
        if (args.ValueKind != JsonValueKind.Array || args.GetArrayLength() == 0)
        {
            var v = GetSetting("countReadAsTreated");
            return new { success = true, exists = v != null, value = v == "1" };
        }
        var a0 = args[0];
        bool value;
        if (a0.ValueKind == JsonValueKind.True) value = true;
        else if (a0.ValueKind == JsonValueKind.False) value = false;
        else if (a0.ValueKind == JsonValueKind.Object && a0.TryGetProperty("value", out var v2))
        {
            if (v2.ValueKind == JsonValueKind.True) value = true;
            else if (v2.ValueKind == JsonValueKind.False) value = false;
            else return new { success = false, error = "value booléen attendu" };
        }
        else return new { success = false, error = "booléen attendu" };
        SetSetting("countReadAsTreated", value ? "1" : "0");
        return new { success = true, value };
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
            return new { success = true, path = dir };
        }
        catch (Exception ex) { return new { success = false, error = ex.Message }; }
    }

    private static object ToLegacyLog(LogEntry e) => new
    {
        timestamp = DateTimeOffset.FromUnixTimeMilliseconds(e.Ts).ToLocalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffzzz"),
        level = e.Level.ToString().ToUpperInvariant(),
        category = e.Category,
        message = e.Message,
        meta = e.Meta,
        data = e.Meta
    };

    private List<object> FilterLogHistory(JsonElement args)
    {
        string level = "ALL", category = "ALL", search = "";
        int limit = 1000;
        if (args.ValueKind == JsonValueKind.Array && args.GetArrayLength() > 0 && args[0].ValueKind == JsonValueKind.Object)
        {
            var f = args[0];
            if (f.TryGetProperty("level", out var l) && l.ValueKind == JsonValueKind.String) level = l.GetString() ?? "ALL";
            if (f.TryGetProperty("category", out var c2) && c2.ValueKind == JsonValueKind.String) category = c2.GetString() ?? "ALL";
            if (f.TryGetProperty("search", out var s) && s.ValueKind == JsonValueKind.String) search = s.GetString() ?? "";
            if (f.TryGetProperty("limit", out var lim) && lim.ValueKind == JsonValueKind.Number) limit = lim.GetInt32();
        }
        var ranks = new Dictionary<string, int> { ["DEBUG"] = 0, ["INFO"] = 1, ["WARN"] = 2, ["ERROR"] = 3 };
        var minRank = ranks.TryGetValue(level.ToUpperInvariant(), out var r) ? r : 0;
        var snap = _log.Snapshot(Math.Max(limit, 100));
        if (snap.Count == 0)
            return ReadLogFileHistory(level, category, search, limit);
        var search2 = search.ToLowerInvariant();
        var list = new List<object>();
        foreach (var e in snap)
        {
            var lvl = e.Level.ToString().ToUpperInvariant();
            if (level.ToUpperInvariant() != "ALL" && (ranks.TryGetValue(lvl, out var er) ? er : -1) < minRank) continue;
            if (category.ToUpperInvariant() != "ALL" && !string.Equals(e.Category, category, StringComparison.OrdinalIgnoreCase)) continue;
            if (search2.Length > 0 && !(e.Message?.ToLowerInvariant().Contains(search2) == true || e.Category.ToLowerInvariant().Contains(search2))) continue;
            list.Add(ToLegacyLog(e));
        }
        return list;
    }

    private List<object> ReadLogFileHistory(string level, string category, string search, int limit)
    {
        var list = new List<object>();
        try
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMonitor", "logs");
            var file = new DirectoryInfo(dir).GetFiles("mailmonitor-*.log").OrderByDescending(f => f.LastWriteTimeUtc).FirstOrDefault();
            if (file is null) return list;

            var ranks = new Dictionary<string, int> { ["DEBUG"] = 0, ["INFO"] = 1, ["WARN"] = 2, ["ERROR"] = 3 };
            var minRank = ranks.TryGetValue((level ?? "ALL").ToUpperInvariant(), out var r) ? r : 0;
            var search2 = (search ?? string.Empty).ToLowerInvariant();
            var lines = File.ReadLines(file.FullName).TakeLast(Math.Max(100, Math.Min(limit, 5000)));

            foreach (var line in lines)
            {
                // Format attendu: yyyy-MM-dd HH:mm:ss.fff [LEVEL] CATEGORY: MESSAGE | META
                var match = System.Text.RegularExpressions.Regex.Match(line, @"^(?<ts>\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?)\s+\[(?<lvl>[A-Z]+)\]\s+(?<cat>[^:]+):\s*(?<msg>.*)$");
                if (!match.Success) continue;

                var lvl = match.Groups["lvl"].Value;
                var cat = match.Groups["cat"].Value.Trim();
                var msgRaw = match.Groups["msg"].Value;
                string msg = msgRaw;
                string? meta = null;
                var sep = msgRaw.IndexOf(" | ", StringComparison.Ordinal);
                if (sep >= 0)
                {
                    msg = msgRaw[..sep].TrimEnd();
                    meta = msgRaw[(sep + 3)..].Trim();
                }

                if ((level ?? "ALL").ToUpperInvariant() != "ALL" && (ranks.TryGetValue(lvl, out var er) ? er : -1) < minRank) continue;
                if ((category ?? "ALL").ToUpperInvariant() != "ALL" && !string.Equals(cat, category, StringComparison.OrdinalIgnoreCase)) continue;

                if (search2.Length > 0)
                {
                    var hay = $"{msg}\n{meta}\n{cat}".ToLowerInvariant();
                    if (!hay.Contains(search2)) continue;
                }

                var ts = match.Groups["ts"].Value;
                if (!DateTime.TryParse(ts, out var parsed)) parsed = DateTime.Now;
                list.Add(new
                {
                    timestamp = new DateTimeOffset(parsed).ToString("yyyy-MM-ddTHH:mm:ss.fffzzz"),
                    level = lvl,
                    category = cat,
                    message = msg,
                    meta,
                    data = meta
                });
            }
        }
        catch { }
        return list;
    }

    private object ExportLogHistory(JsonElement args)
    {
        try
        {
            var entries = FilterLogHistory(args);
            string? dest = null;
            _ownerForm.Invoke(() =>
            {
                using var dlg = new SaveFileDialog
                {
                    Title = "Exporter l'historique des logs",
                    Filter = "Fichier texte (*.log)|*.log|Tous (*.*)|*.*",
                    FileName = $"mailmonitor-logs-{DateTime.Now:yyyyMMdd-HHmmss}.log"
                };
                if (dlg.ShowDialog(_ownerForm) == DialogResult.OK) dest = dlg.FileName;
            });
            if (string.IsNullOrEmpty(dest)) return new { success = false, canceled = true };
            var sb = new System.Text.StringBuilder();
            foreach (var e in _log.Snapshot(2000))
            {
                var ts = DateTimeOffset.FromUnixTimeMilliseconds(e.Ts).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
                sb.AppendLine($"{ts} [{e.Level.ToString().ToUpperInvariant()}] {e.Category}: {e.Message}");
            }
            File.WriteAllText(dest!, sb.ToString());
            return new { success = true, exported = entries.Count, path = dest };
        }
        catch (Exception ex) { return new { success = false, error = ex.Message }; }
    }
}
