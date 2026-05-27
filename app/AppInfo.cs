using System.Reflection;

namespace MailMonitor;

public static class AppInfo
{
    public static string Version =>
        Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.0.0";

    public const string ProductName = "Mail Monitor";
    public const string GitHubOwner = "Erkaek";
    public const string GitHubRepo = "AutoMailMonitor";
}

public sealed class AppPaths
{
    public string Root { get; private init; } = "";
    public string DataDir { get; private init; } = "";
    public string LogsDir { get; private init; } = "";
    public string UpdatesDir { get; private init; } = "";
    public string WwwRoot { get; private init; } = "";
    public string WebView2UserData { get; private init; } = "";
    public string DatabasePath { get; private init; } = "";

    public static AppPaths Initialize()
    {
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var root = Path.Combine(local, "MailMonitor");
        var data = Path.Combine(root, "data");
        var logs = Path.Combine(root, "logs");
        var upd  = Path.Combine(root, "updates");
        var wv2  = Path.Combine(root, "webview2");
        Directory.CreateDirectory(data);
        Directory.CreateDirectory(logs);
        Directory.CreateDirectory(upd);
        Directory.CreateDirectory(wv2);

        var exeDir = AppContext.BaseDirectory;
        var www = Path.Combine(exeDir, "wwwroot");
        if (!Directory.Exists(www) || !File.Exists(Path.Combine(www, "index.html")))
        {
            var devPublic = Path.GetFullPath(Path.Combine(exeDir, "..", "..", "..", "..", "public"));
            if (Directory.Exists(devPublic)) www = devPublic;
        }

        return new AppPaths
        {
            Root = root,
            DataDir = data,
            LogsDir = logs,
            UpdatesDir = upd,
            WwwRoot = www,
            WebView2UserData = wv2,
            DatabasePath = Path.Combine(data, "mailmonitor.db")
        };
    }
}
