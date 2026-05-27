using System.Collections.Concurrent;
using System.Runtime.InteropServices;

namespace MailMonitor.Services;

public sealed class OutlookFolderRef
{
    public string StoreId { get; init; } = "";
    public string EntryId { get; init; } = "";
    public string Path    { get; init; } = "";
    public string Name    { get; init; } = "";
    public int    ItemCount { get; init; }
}

public sealed class OutlookMailItem
{
    public string EntryId { get; init; } = "";
    public string Subject { get; init; } = "";
    public string Sender  { get; init; } = "";
    public DateTime ReceivedTime { get; init; }
    public bool IsUnread { get; init; }
}

public sealed class OutlookEvent
{
    public enum Kind { Add, Change, Remove }
    public Kind  Type { get; init; }
    public string FolderEntryId { get; init; } = "";
    public string StoreId { get; init; } = "";
    public string MailEntryId { get; init; } = "";
    public OutlookMailItem? Mail { get; init; }
}

/// <summary>
/// Late-binding intégral via ProgID + IDispatch (dynamic). Zéro dépendance PIA,
/// fonctionne sur n'importe quelle version d'Outlook installée (2013→365),
/// sans office.dll, sans droits admin. Pump COM sur STA dédié.
/// </summary>
public sealed class OutlookService : IDisposable
{
    private readonly LogService _log;
    private readonly Thread _staThread;
    private readonly BlockingCollection<Action> _queue = new(boundedCapacity: 1024);
    private readonly CancellationTokenSource _cts = new();
    private readonly TaskCompletionSource _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);

    private dynamic? _app;
    private dynamic? _ns;
    public bool IsConnected { get; private set; }

    public OutlookService(LogService log)
    {
        _log = log;
        _staThread = new Thread(StaPump) { IsBackground = true, Name = "OutlookSTA" };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();
    }

    public Task WhenReady => _ready.Task;

    private void StaPump()
    {
        try { ConnectInternal(); IsConnected = true; _ready.TrySetResult(); }
        catch (Exception ex)
        {
            _log.Error("OUTLOOK", "Connexion COM échouée", ex);
            _ready.TrySetException(ex);
            return;
        }

        foreach (var action in _queue.GetConsumingEnumerable(_cts.Token))
        {
            try { action(); }
            catch (Exception ex) { _log.Error("OUTLOOK", "Action COM échouée", ex); }
        }
    }

    private void ConnectInternal()
    {
        var t = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Outlook n'est pas installé sur ce poste.");
        _app = Activator.CreateInstance(t)!;
        _ns = _app!.GetNamespace("MAPI");
        try { _ns!.Logon(Type.Missing, Type.Missing, false, false); }
        catch { /* session existante OK */ }
        _log.Info("OUTLOOK", "Connexion COM établie (late-binding)");
    }

    public Task<T> InvokeAsync<T>(Func<dynamic, dynamic, T> fn)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        Action action = () =>
        {
            try { tcs.SetResult(fn(_app!, _ns!)); }
            catch (Exception ex) { tcs.SetException(ex); }
        };
        try
        {
            if (!_queue.TryAdd(action, 5000, _cts.Token))
                tcs.TrySetException(new TimeoutException("Le délai d'attente pour planifier l'action Outlook a expiré car la file COM est saturée."));
        }
        catch (OperationCanceledException ex)
        {
            tcs.TrySetException(ex);
        }
        catch (InvalidOperationException ex)
        {
            tcs.TrySetException(ex);
        }
        return tcs.Task;
    }

    public Task InvokeAsync(Action<dynamic, dynamic> fn) =>
        InvokeAsync<object?>((a, n) => { fn(a, n); return null!; });

    public Task<List<(string id, string name)>> ListStoresAsync() => InvokeAsync<List<(string, string)>>((_, ns) =>
    {
        var list = new List<(string, string)>();
        dynamic stores = ns.Stores;
        int count = (int)stores.Count;
        for (int i = 1; i <= count; i++)
        {
            try
            {
                dynamic st = stores.Item(i);
                list.Add(((string)st.StoreID, (string)st.DisplayName));
            }
            catch { }
        }
        return list;
    });

    public Task<List<OutlookFolderRef>> ListFoldersAsync(string storeId, int maxDepth = 8) =>
        InvokeAsync<List<OutlookFolderRef>>((_, ns) =>
    {
        var result = new List<OutlookFolderRef>(128);
        dynamic? store = null;
        dynamic stores = ns.Stores;
        int sc = (int)stores.Count;
        for (int i = 1; i <= sc; i++)
        {
            dynamic st = stores.Item(i);
            if ((string)st.StoreID == storeId) { store = st; break; }
        }
        if (store is null) return result;

        dynamic root = store.GetRootFolder();
        Walk(root, "", 0);
        return result;

        void Walk(dynamic f, string parentPath, int depth)
        {
            string name = (string)f.Name;
            string path = string.IsNullOrEmpty(parentPath) ? name : parentPath + "/" + name;
            try
            {
                result.Add(new OutlookFolderRef
                {
                    StoreId = storeId,
                    EntryId = (string)f.EntryID,
                    Path = path,
                    Name = name,
                    ItemCount = SafeCount(f)
                });
            }
            catch { }
            if (depth >= maxDepth) return;
            try
            {
                dynamic subs = f.Folders;
                int cnt = (int)subs.Count;
                for (int i = 1; i <= cnt; i++)
                {
                    try { Walk(subs.Item(i), path, depth + 1); } catch { }
                }
            }
            catch { }
        }
        static int SafeCount(dynamic f) { try { return (int)f.Items.Count; } catch { return 0; } }
    });

    public Task<List<OutlookMailItem>> ScanFolderAsync(string storeId, string folderEntryId, DateTime? since = null) =>
        InvokeAsync<List<OutlookMailItem>>((_, ns) =>
        {
            var list = new List<OutlookMailItem>(512);
            dynamic folder;
            try { folder = ns.GetFolderFromID(folderEntryId, storeId); }
            catch (Exception ex) { _log.Warn("OUTLOOK", "GetFolderFromID KO: " + ex.Message); return list; }

            dynamic items = folder.Items;
            try { items.SetColumns("EntryID,Subject,ReceivedTime,SenderName,UnRead"); } catch { }

            if (since.HasValue)
            {
                try
                {
                    string filter = "[ReceivedTime] >= '" + since.Value.ToString("g") + "'";
                    items = items.Restrict(filter);
                    try { items.SetColumns("EntryID,Subject,ReceivedTime,SenderName,UnRead"); } catch { }
                }
                catch { }
            }
            try { items.Sort("[ReceivedTime]", true); } catch { }

            int count;
            try { count = (int)items.Count; } catch { return list; }
            for (int i = 1; i <= count; i++)
            {
                dynamic? mi = null;
                try { mi = items[i]; } catch { continue; }
                try
                {
                    string? cls = SafeStr(() => (string)mi.MessageClass);
                    if (cls is not null && !cls.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase)) continue;
                    list.Add(new OutlookMailItem
                    {
                        EntryId = SafeStr(() => (string)mi.EntryID) ?? "",
                        Subject = SafeStr(() => (string)mi.Subject) ?? "",
                        Sender  = SafeStr(() => (string)mi.SenderName) ?? "",
                        ReceivedTime = SafeDt(() => (DateTime)mi.ReceivedTime),
                        IsUnread = SafeBool(() => (bool)mi.UnRead)
                    });
                }
                catch { }
                finally { try { Marshal.ReleaseComObject((object)mi!); } catch { } }
            }
            return list;
        });

    private static string? SafeStr(Func<string> f) { try { return f(); } catch { return null; } }
    private static DateTime SafeDt(Func<DateTime> f) { try { return f(); } catch { return DateTime.MinValue; } }
    private static bool SafeBool(Func<bool> f) { try { return f(); } catch { return false; } }

    public void Dispose()
    {
        try { _queue.Add(() => { try { _ns?.Logoff(); } catch { } }); } catch { }
        try { _queue.CompleteAdding(); } catch { }
        bool stopped = false;
        try { stopped = _staThread.Join(TimeSpan.FromSeconds(2)); } catch { }
        if (!stopped)
        {
            try { _cts.Cancel(); } catch { }
            try { _staThread.Join(TimeSpan.FromSeconds(2)); } catch { }
        }
        try { if (_app is not null) Marshal.FinalReleaseComObject((object)_app); } catch { }
    }
}
