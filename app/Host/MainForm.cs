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
        try
        {
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
}
