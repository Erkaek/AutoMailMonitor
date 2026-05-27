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

    private async Task<object?> DispatchAsync(string method, JsonElement args)
    {
        switch (method)
        {
            case "app.version": return AppInfo.Version;
            case "app.autostart.get": return _autostart.IsEnabled;
            case "app.autostart.set":
                if (args.ValueKind == JsonValueKind.Array && args[0].GetBoolean()) _autostart.EnsureEnabled();
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
                    var storeId = args[0].GetString() ?? "";
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
                    var p = args[0];
                    await _monitor.AddFolderAsync(
                        p.GetProperty("storeId").GetString()!,
                        p.GetProperty("entryId").GetString()!,
                        p.GetProperty("path").GetString()!,
                        p.GetProperty("displayName").GetString() ?? "",
                        p.TryGetProperty("category", out var c) ? c.GetString() : null);
                    return null;
                }
            case "folders.remove":
                await _monitor.RemoveFolderAsync(args[0].GetString()!);
                return null;

            case "stats.summary": return BuildSummary();
            case "stats.weekly": return BuildWeekly(args);
            case "stats.by-category": return BuildByCategory();

            case "emails.recent": return RecentEmails(args);

            case "logs.recent": return _log.Snapshot(500);

            case "weekly-comments.list":
                return ListWeeklyComments(args[0].GetInt32(), args[1].GetInt32());
            case "weekly-comments.add":
                return AddWeeklyComment(args[0]);

            case "window.minimize": _ownerForm.BeginInvoke(new Action(() => _ownerForm.WindowState = FormWindowState.Minimized)); return null;
            case "window.close":    _ownerForm.BeginInvoke(new Action(() => _ownerForm.Hide())); return null;

            default: throw new InvalidOperationException("Méthode inconnue: " + method);
        }
    }

    private List<object> ListMonitoredFolders()
    {
        var list = new List<object>();
        using var cmd = _storage.Connection.CreateCommand();
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
        using var cmd = _storage.Connection.CreateCommand();
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
        using var cmd = _storage.Connection.CreateCommand();
        cmd.CommandText = "SELECT category, COUNT(*) FROM emails GROUP BY category";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(new { category = r.GetString(0), count = r.GetInt64(1) });
        return list;
    }

    private List<object> BuildWeekly(JsonElement args)
    {
        var weeks = args.ValueKind == JsonValueKind.Array && args.GetArrayLength() > 0 ? args[0].GetInt32() : 12;
        var list = new List<object>();
        using var cmd = _storage.Connection.CreateCommand();
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
        using var cmd = _storage.Connection.CreateCommand();
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
        using var cmd = _storage.Connection.CreateCommand();
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
            using var cmd = _storage.Connection.CreateCommand();
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
}
