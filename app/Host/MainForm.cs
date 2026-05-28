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
        //   1) Si un runtime Edge WebView2 est déjà disponible (Evergreen système OU
        //      install per-user précédente dans %LocalAppData%\Microsoft\EdgeWebView)
        //      → on l'utilise.
        //   2) Sinon, on extrait le bootstrapper MicrosoftEdgeWebview2Setup.exe
        //      embarqué dans l'exe et on le lance en mode silencieux per-user
        //      (--install --webview --runtime --msedge --user-level-install). Pas
        //      de droits admin requis, install dans %LocalAppData%.
        string? installedVersion = null;
        try { installedVersion = CoreWebView2Environment.GetAvailableBrowserVersionString(); }
        catch { installedVersion = null; }

        if (string.IsNullOrEmpty(installedVersion))
        {
            _log.Info("WEBVIEW", "Runtime absent → lancement bootstrapper per-user…");
            var installed = await TryInstallEvergreenPerUserAsync();
            if (!installed)
            {
                _log.Error("WEBVIEW", "Installation auto du runtime KO.");
                MessageBox.Show(
                    "WebView2 indisponible : impossible d'installer automatiquement le runtime Edge.\r\n\r\n" +
                    "Téléchargez et installez manuellement (sans droits admin) :\r\n" +
                    "https://go.microsoft.com/fwlink/p/?LinkId=2124703",
                    "Mail Monitor", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }
            try { installedVersion = CoreWebView2Environment.GetAvailableBrowserVersionString(); }
            catch { installedVersion = null; }
            if (string.IsNullOrEmpty(installedVersion))
            {
                _log.Error("WEBVIEW", "Runtime toujours introuvable après install.");
                MessageBox.Show("WebView2 installé mais introuvable. Redémarrez l'application.",
                    "Mail Monitor", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }
        }
        _log.Info("WEBVIEW", "Runtime détecté: " + installedVersion);

        try
        {
            try { Directory.CreateDirectory(_paths.WebView2UserData); }
            catch (Exception ex)
            {
                _log.Error("WEBVIEW", "Création UserDataFolder KO: " + _paths.WebView2UserData, ex);
                throw;
            }

            var env = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: _paths.WebView2UserData);

            await _web.EnsureCoreWebView2Async(env);

            _web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app.local", _paths.WwwRoot, CoreWebView2HostResourceAccessKind.Allow);

            // Sécurité : restreindre strictement la navigation à app.local (review Copilot)
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
        }
        catch (Exception ex)
        {
            _log.Error("WEBVIEW", "Init KO", ex);
            MessageBox.Show("WebView2 indisponible : " + ex.Message, "Mail Monitor",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            using var stream = asm.GetManifestResourceStream("MicrosoftEdgeWebview2Setup.exe");
            if (stream is null)
            {
                _log.Error("WEBVIEW", "Ressource bootstrapper introuvable dans l'exe.");
                return false;
            }
            var tmp = Path.Combine(Path.GetTempPath(), "MailMonitor-WebView2Setup-" + Guid.NewGuid().ToString("N") + ".exe");
            using (var fs = File.Create(tmp)) await stream.CopyToAsync(fs);

            // Progress UI minimal pour ne pas laisser l'utilisateur sans feedback (DL peut prendre 1-2 min)
            var splash = new Form
            {
                Text = "Mail Monitor — Première installation",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false, ControlBox = false,
                ClientSize = new Size(480, 110), TopMost = true
            };
            var lbl = new Label
            {
                Text = "Installation du runtime Edge WebView2 (per-user, sans droits admin)…\r\n" +
                       "Cette étape ne se produit qu'au premier lancement.",
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
                    // Args silencieux per-user documentés Microsoft pour le bootstrapper Evergreen.
                    Arguments = "/silent /install",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var proc = Process.Start(psi);
                if (proc is null) return false;
                await Task.Run(() => proc.WaitForExit(5 * 60 * 1000));
                _log.Info("WEBVIEW", "Bootstrapper exit code: " + proc.ExitCode);
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
