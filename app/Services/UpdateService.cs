using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json.Serialization;
using System.Diagnostics;

namespace MailMonitor.Services;

public sealed class UpdateService
{
    private readonly AppPaths _paths;
    private readonly LogService _log;
    private static readonly HttpClient Http = CreateHttp();

    public event Action<string>? OnAvailable;
    public event Action<int>? OnProgress;
    public event Action<string>? OnError;
    public event Action? OnReadyToInstall;

    public UpdateService(AppPaths paths, LogService log) { _paths = paths; _log = log; }

    private static HttpClient CreateHttp()
    {
        var h = new HttpClient(new SocketsHttpHandler
        {
            ConnectTimeout = TimeSpan.FromSeconds(15),
            PooledConnectionLifetime = TimeSpan.FromMinutes(5)
        })
        {
            Timeout = TimeSpan.FromMinutes(5)
        };
        h.DefaultRequestHeaders.UserAgent.ParseAdd("MailMonitor/" + AppInfo.Version);
        h.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return h;
    }

    public async Task CheckOnceAsync()
    {
        try
        {
            var url = $"https://api.github.com/repos/{AppInfo.GitHubOwner}/{AppInfo.GitHubRepo}/releases/latest";
            var rel = await Http.GetFromJsonAsync<Release>(url);
            if (rel?.TagName is null) return;
            var latest = rel.TagName.TrimStart('v');
            if (!IsNewer(latest, AppInfo.Version)) { _log.Info("UPDATE", $"À jour ({AppInfo.Version})"); return; }
            _log.Info("UPDATE", $"Nouvelle version dispo: {latest}");
            OnAvailable?.Invoke(latest);

            var exeAsset = rel.Assets?.FirstOrDefault(a =>
                a.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) &&
                !a.Name.Contains("setup", StringComparison.OrdinalIgnoreCase));
            if (exeAsset is null) { _log.Warn("UPDATE", "Asset .exe introuvable"); return; }
            await DownloadAsync(exeAsset.BrowserDownloadUrl, exeAsset.Name);
            OnReadyToInstall?.Invoke();
        }
        catch (Exception ex) { _log.Warn("UPDATE", "Check KO: " + ex.Message); OnError?.Invoke(ex.Message); }
    }

    private async Task DownloadAsync(string url, string fileName)
    {
        var dest = Path.Combine(_paths.UpdatesDir, fileName);
        using var resp = await Http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
        resp.EnsureSuccessStatusCode();
        var total = resp.Content.Headers.ContentLength ?? -1L;
        await using var src = await resp.Content.ReadAsStreamAsync();
        await using var dst = File.Create(dest);
        var buf = new byte[81920]; long got = 0; int read; int lastPct = -1;
        while ((read = await src.ReadAsync(buf)) > 0)
        {
            await dst.WriteAsync(buf.AsMemory(0, read));
            got += read;
            if (total > 0)
            {
                var pct = (int)(got * 100 / total);
                if (pct != lastPct) { lastPct = pct; OnProgress?.Invoke(pct); }
            }
        }
        _log.Info("UPDATE", $"Téléchargé: {dest} ({got} octets)");
        var pending = Path.Combine(_paths.UpdatesDir, "pending.json");
        await File.WriteAllTextAsync(pending,
            $"{{\"file\":\"{dest.Replace("\\", "\\\\")}\"}}");
    }

    private static bool IsNewer(string remote, string local)
    {
        try { return new Version(remote) > new Version(local); }
        catch { return !string.Equals(remote, local, StringComparison.OrdinalIgnoreCase); }
    }

    public void ApplyAndRestart()
    {
        try
        {
            var pending = Path.Combine(_paths.UpdatesDir, "pending.json");
            if (!File.Exists(pending)) return;
            var json = File.ReadAllText(pending);
            var match = System.Text.RegularExpressions.Regex.Match(json, "\"file\"\\s*:\\s*\"([^\"]+)\"");
            if (!match.Success) return;
            var newExe = match.Groups[1].Value.Replace("\\\\", "\\");
            var currentExe = Process.GetCurrentProcess().MainModule!.FileName!;
            var bat = Path.Combine(_paths.UpdatesDir, "swap.bat");
            File.WriteAllText(bat, $"""
@echo off
:loop
tasklist /FI "PID eq {Environment.ProcessId}" | find "{Environment.ProcessId}" >NUL
if not errorlevel 1 ( ping 127.0.0.1 -n 2 >NUL & goto loop )
copy /Y "{newExe}" "{currentExe}"
del "{pending}"
start "" "{currentExe}" --minimized
del "%~f0"
""");
            Process.Start(new ProcessStartInfo("cmd.exe", $"/c \"{bat}\"")
            {
                CreateNoWindow = true, UseShellExecute = false, WindowStyle = ProcessWindowStyle.Hidden
            });
            Application.Exit();
        }
        catch (Exception ex) { _log.Error("UPDATE", "Apply KO", ex); }
    }

    public sealed class Release
    {
        [JsonPropertyName("tag_name")] public string? TagName { get; set; }
        [JsonPropertyName("assets")]   public List<Asset>? Assets { get; set; }
    }
    public sealed class Asset
    {
        [JsonPropertyName("name")] public string Name { get; set; } = "";
        [JsonPropertyName("browser_download_url")] public string BrowserDownloadUrl { get; set; } = "";
    }
}
