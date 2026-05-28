using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using MailMonitor.Services;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace MailMonitor.Host;

public sealed class MainForm : Form
{
    private readonly WebView2 _web = new() { Dock = DockStyle.Fill };
    private readonly NotifyIcon _tray = new();
    private readonly MonitoringService _monitor;
    private readonly StorageService _storage;
    private readonly LogService _log;
    private readonly UpdateService _updater;
    private readonly AutoStartService _autostart;
    private readonly AppPaths _paths;
    private WebBridge? _bridge;

    public MainForm(MonitoringService monitor, StorageService storage, LogService log,
                    UpdateService updater, AutoStartService autostart, AppPaths paths, bool startMinimized)
    {
        _monitor = monitor; _storage = storage; _log = log;
        _updater = updater; _autostart = autostart; _paths = paths;

        Text = "Mail Monitor";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(1280, 800);
        MinimumSize = new Size(900, 600);
        Icon = SystemIcons.Application;
        Controls.Add(_web);

        SetupTray();
        if (startMinimized) { WindowState = FormWindowState.Minimized; ShowInTaskbar = false; }
        Shown += async (_, _) => await InitWebViewAsync();
        FormClosing += (s, e) =>
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true; Hide(); ShowInTaskbar = false;
            }
        };
    }

    private void SetupTray()
    {
        _tray.Text = "Mail Monitor";
        _tray.Icon = SystemIcons.Application;
        _tray.Visible = true;
        var menu = new ContextMenuStrip();
        menu.Items.Add("Ouvrir", null, (_, _) => RestoreFromTray());
        menu.Items.Add("Vérifier les mises à jour", null, async (_, _) => await _updater.CheckOnceAsync());
        menu.Items.Add(new ToolStripSeparator());
        var autoItem = new ToolStripMenuItem("Lancer au démarrage") { Checked = _autostart.IsEnabled, CheckOnClick = true };
        autoItem.Click += (_, _) => { if (autoItem.Checked) _autostart.EnsureEnabled(); else _autostart.Disable(); };
        menu.Items.Add(autoItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Quitter", null, (_, _) => { _tray.Visible = false; Application.Exit(); });
        _tray.ContextMenuStrip = menu;
        _tray.DoubleClick += (_, _) => RestoreFromTray();
    }

    private void RestoreFromTray()
    {
        Show(); WindowState = FormWindowState.Normal; ShowInTaskbar = true; Activate();
    }

    private async Task InitWebViewAsync()
    {
        // Stratégie zéro-install pour PC verrouillé :
        //   1) On TENTE directement CreateAsync. Si ça marche → fin.
        //   2) Sinon (runtime absent OU registry pointant vers un dossier inexistant
        //      → erreur 0x80040003), on exécute le bootstrapper embarqué en mode
        //      silencieux per-user (zéro droit admin, install %LocalAppData%).
        //   3) On retente CreateAsync.
        try { Directory.CreateDirectory(_paths.WebView2UserData); }
        catch (Exception ex)
        {
            _log.Error("WEBVIEW", "Création UserDataFolder KO: " + _paths.WebView2UserData, ex);
        }

        var firstError = await TryInitOnceAsync();
        if (firstError is null) return; // OK

        _log.Warn("WEBVIEW", "1er essai KO (" + firstError + ") → tentative install bootstrapper…");
        var installed = await TryInstallEvergreenPerUserAsync();
        if (!installed)
        {
            _log.Error("WEBVIEW", "Installation auto du runtime KO.");
            MessageBox.Show(
                "WebView2 indisponible : impossible d'installer automatiquement le runtime Edge.\r\n\r\n" +
                "Détail : " + firstError + "\r\n\r\n" +
                "Téléchargez et installez manuellement (sans droits admin) :\r\n" +
                "https://go.microsoft.com/fwlink/p/?LinkId=2124703",
                "Mail Monitor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Application.Exit();
            return;
        }

        var secondError = await TryInitOnceAsync();
        if (secondError is null) return; // OK après install

        _log.Error("WEBVIEW", "Init KO après install: " + secondError);
        MessageBox.Show(
            "WebView2 installé mais l'initialisation a échoué.\r\n\r\n" +
            "Détail : " + secondError + "\r\n\r\n" +
            "Essayez de relancer Mail Monitor. Si l'erreur persiste, redémarrez la session Windows.",
            "Mail Monitor", MessageBoxButtons.OK, MessageBoxIcon.Error);
        Application.Exit();
    }

    /// <summary>
    /// Tente une init WebView2 complète. Renvoie null si OK, sinon le message d'erreur.
    /// </summary>
    private async Task<string?> TryInitOnceAsync()
    {
        string? version = null;
        try { version = CoreWebView2Environment.GetAvailableBrowserVersionString(); }
        catch (Exception ex) { _log.Warn("WEBVIEW", "GetAvailableBrowserVersionString KO: " + ex.Message); }

        if (string.IsNullOrEmpty(version))
            return "Runtime Edge WebView2 introuvable (GetAvailableBrowserVersionString=null).";

        _log.Info("WEBVIEW", "Version détectée: " + version);

        try
        {
            var env = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: _paths.WebView2UserData);

            await _web.EnsureCoreWebView2Async(env);

            _web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app.local", _paths.WwwRoot, CoreWebView2HostResourceAccessKind.Allow);

            _web.CoreWebView2.NavigationStarting += (_, e) =>
            {
                if (!string.IsNullOrEmpty(e.Uri) &&
                    !e.Uri.StartsWith("https://app.local/", StringComparison.OrdinalIgnoreCase) &&
                    !e.Uri.StartsWith("about:", StringComparison.OrdinalIgnoreCase))
                {
                    e.Cancel = true;
                    _log.Warn("WEBVIEW", "Navigation bloquée vers " + e.Uri);
                }
            };
            _web.CoreWebView2.NewWindowRequested += (_, e) => { e.Handled = true; };

#if !DEBUG
            _web.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _web.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _web.CoreWebView2.Settings.IsZoomControlEnabled = false;
#endif
            _web.CoreWebView2.Settings.IsStatusBarEnabled = false;

            _bridge = new WebBridge(_web, _monitor, _storage, _log, _updater, _autostart, this);
            _bridge.Attach();

            _web.CoreWebView2.Navigate("https://app.local/index.html");
            return null;
        }
        catch (Exception ex)
        {
            _log.Error("WEBVIEW", "CreateAsync/EnsureCoreWebView2 KO", ex);
            // Cas typique sur PC d'entreprise : registry pointe vers un dossier purgé.
            // → renvoie l'erreur pour permettre au caller de tenter le bootstrapper.
            return ex.Message + (ex.HResult != 0 ? $" (HRESULT 0x{ex.HResult:X8})" : "");
        }
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _tray.Visible = false; _tray.Dispose();
        base.OnFormClosed(e);
    }

    /// <summary>
    /// Extrait le bootstrapper Edge WebView2 embarqué et l'exécute en mode
    /// per-user silencieux (zéro droit admin, install dans %LocalAppData%).
    /// </summary>
    private async Task<bool> TryInstallEvergreenPerUserAsync()
    {
        try
        {
            var asm = Assembly.GetExecutingAssembly();

            // Log de toutes les ressources embarquées pour diagnostiquer un éventuel
            // problème de nommage si le bootstrapper n'a pas été embarqué par le CI.
            var allRes = asm.GetManifestResourceNames();
            _log.Info("WEBVIEW", "Ressources embarquées: " + string.Join(", ", allRes));

            var resName = allRes.FirstOrDefault(n =>
                n.Equals("MicrosoftEdgeWebview2Setup.exe", StringComparison.OrdinalIgnoreCase) ||
                n.EndsWith(".MicrosoftEdgeWebview2Setup.exe", StringComparison.OrdinalIgnoreCase));

            using var stream = resName is null ? null : asm.GetManifestResourceStream(resName);
            if (stream is null)
            {
                _log.Error("WEBVIEW", "Bootstrapper Edge non embarqué dans l'exe (ressource introuvable).");
                return false;
            }
            var tmp = Path.Combine(Path.GetTempPath(), "MailMonitor-WebView2Setup-" + Guid.NewGuid().ToString("N") + ".exe");
            using (var fs = File.Create(tmp)) await stream.CopyToAsync(fs);
            _log.Info("WEBVIEW", $"Bootstrapper extrait: {tmp} ({new FileInfo(tmp).Length / 1024} KB)");

            var splash = new Form
            {
                Text = "Mail Monitor — Première installation",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false, ControlBox = false,
                ClientSize = new Size(520, 130), TopMost = true
            };
            var lbl = new Label
            {
                Text = "Installation du runtime Edge WebView2…\r\n" +
                       "Mode per-user, sans droits administrateur. Cette étape ne se produit\r\n" +
                       "qu'au premier lancement et peut prendre 1 à 2 minutes.",
                AutoSize = false, Dock = DockStyle.Fill, Padding = new Padding(16),
                TextAlign = ContentAlignment.MiddleCenter
            };
            splash.Controls.Add(lbl);
            splash.Show(this);
            Application.DoEvents();

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = tmp,
                    Arguments = "/silent /install",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                };
                using var proc = Process.Start(psi);
                if (proc is null)
                {
                    _log.Error("WEBVIEW", "Process.Start a renvoyé null pour le bootstrapper.");
                    return false;
                }
                var stdOutTask = proc.StandardOutput.ReadToEndAsync();
                var stdErrTask = proc.StandardError.ReadToEndAsync();
                var exited = await Task.Run(() => proc.WaitForExit(5 * 60 * 1000));
                var stdOut = await stdOutTask; var stdErr = await stdErrTask;
                if (!exited)
                {
                    _log.Error("WEBVIEW", "Bootstrapper timeout (5 min).");
                    try { proc.Kill(); } catch { }
                    return false;
                }
                _log.Info("WEBVIEW", $"Bootstrapper exit={proc.ExitCode}");
                if (!string.IsNullOrWhiteSpace(stdOut)) _log.Info("WEBVIEW", "STDOUT: " + stdOut.Trim());
                if (!string.IsNullOrWhiteSpace(stdErr)) _log.Warn("WEBVIEW", "STDERR: " + stdErr.Trim());

                // L'installer Edge peut renvoyer 0 (succès) ou des codes positifs valides
                // (ex: 17 = déjà installé). On considère exit != 0 comme un échec mais on
                // tentera quand même un retry d'init au cas où.
                return proc.ExitCode == 0;
            }
            finally
            {
                splash.Close(); splash.Dispose();
                try { File.Delete(tmp); } catch { }
            }
        }
        catch (Exception ex)
        {
            _log.Error("WEBVIEW", "Bootstrapper KO", ex);
            return false;
        }
    }
}
