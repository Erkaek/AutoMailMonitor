using System.Threading;
using MailMonitor.Host;
using MailMonitor.Services;

namespace MailMonitor;

internal static class Program
{
    private static Mutex? _singleInstance;

    [STAThread]
    private static void Main(string[] args)
    {
        _singleInstance = new Mutex(initiallyOwned: true, name: @"Local\MailMonitor.SingleInstance", out var createdNew);
        if (!createdNew) return;

        ApplicationConfiguration.Initialize();
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);

        bool startMinimized = args.Any(a =>
            string.Equals(a, "--minimized", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(a, "/minimized", StringComparison.OrdinalIgnoreCase));

        var paths = AppPaths.Initialize();
        var log = new LogService(paths);
        log.Info("APP", $"MailMonitor v{AppInfo.Version} démarrage (minimized={startMinimized})");

        var storage = new StorageService(paths, log);
        storage.Initialize();

        var outlook = new OutlookService(log);
        var classification = new ClassificationService(storage, log);
        var monitoring = new MonitoringService(outlook, storage, classification, log);
        var autoStart = new AutoStartService(log);
        var updater = new UpdateService(paths, log);


        try
        {
            _ = Task.Run(async () =>
            {
                try { await monitoring.StartAsync(); }
                catch (Exception ex) { log.Error("MONITOR", "Démarrage monitoring échoué", ex); }
            });

            _ = Task.Run(async () =>
            {
                try { await updater.CheckOnceAsync(); }
                catch (Exception ex) { log.Warn("UPDATE", "Check init échoué: " + ex.Message); }
            });

            using var form = new MainForm(monitoring, storage, log, updater, autoStart, paths, startMinimized);
            Application.Run(form);
        }
        finally
        {
            try { monitoring.Stop(); } catch { }
            try { outlook.Dispose(); } catch { }
            try { storage.Dispose(); } catch { }
            try { log.Info("APP", "Arrêt propre"); log.Dispose(); } catch { }
            _singleInstance?.ReleaseMutex();
            _singleInstance?.Dispose();
        }
    }
}
