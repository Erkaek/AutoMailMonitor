using Microsoft.Win32;
using System.Diagnostics;

namespace MailMonitor.Services;

public sealed class AutoStartService
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "MailMonitor";
    private readonly LogService _log;

    public AutoStartService(LogService log) { _log = log; }

    public bool IsEnabled
    {
        get
        {
            try
            {
                using var k = Registry.CurrentUser.OpenSubKey(RunKey, false);
                return k?.GetValue(ValueName) is string s && !string.IsNullOrWhiteSpace(s);
            }
            catch { return false; }
        }
    }

    public void EnsureEnabled()
    {
        try
        {
            var exe = Process.GetCurrentProcess().MainModule?.FileName;
            if (string.IsNullOrEmpty(exe)) return;
            var cmd = $"\"{exe}\" --minimized";
            using var k = Registry.CurrentUser.CreateSubKey(RunKey, true);
            var current = k.GetValue(ValueName) as string;
            if (current != cmd)
            {
                k.SetValue(ValueName, cmd, RegistryValueKind.String);
                _log.Info("AUTOSTART", "Auto-start activé: " + cmd);
            }
        }
        catch (Exception ex) { _log.Warn("AUTOSTART", "Activation échouée: " + ex.Message); }
    }

    public void Disable()
    {
        try
        {
            using var k = Registry.CurrentUser.OpenSubKey(RunKey, true);
            k?.DeleteValue(ValueName, throwOnMissingValue: false);
            _log.Info("AUTOSTART", "Auto-start désactivé");
        }
        catch (Exception ex) { _log.Warn("AUTOSTART", "Désactivation échouée: " + ex.Message); }
    }
}
