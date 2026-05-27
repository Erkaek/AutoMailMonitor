using System.Collections.Concurrent;
using System.Text;

namespace MailMonitor.Services;

public enum LogLevel { Debug, Info, Warn, Error }

public sealed class LogEntry
{
    public long Ts { get; init; }
    public LogLevel Level { get; init; }
    public string Category { get; init; } = "";
    public string Message { get; init; } = "";
    public string? Meta { get; init; }
}

public sealed class LogService : IDisposable
{
    private readonly AppPaths _paths;
    private readonly BlockingCollection<LogEntry> _queue = new(boundedCapacity: 8192);
    private readonly Thread _worker;
    private readonly CancellationTokenSource _cts = new();
    private readonly ConcurrentQueue<LogEntry> _recent = new();
    private const int RecentCap = 2000;
    private readonly object _fileLock = new();
    private string _currentFile = "";
    public event Action<LogEntry>? OnEntry;

    public LogService(AppPaths paths)
    {
        _paths = paths;
        RollFile();
        _worker = new Thread(WorkerLoop) { IsBackground = true, Name = "LogWriter" };
        _worker.Start();
    }

    public void Debug(string cat, string msg) => Enqueue(LogLevel.Debug, cat, msg);
    public void Info(string cat, string msg)  => Enqueue(LogLevel.Info,  cat, msg);
    public void Warn(string cat, string msg)  => Enqueue(LogLevel.Warn,  cat, msg);
    public void Error(string cat, string msg, Exception? ex = null) =>
        Enqueue(LogLevel.Error, cat, msg, ex?.ToString());

    public IReadOnlyList<LogEntry> Snapshot(int limit = 500) =>
        _recent.ToArray().TakeLast(limit).ToList();

    private void Enqueue(LogLevel level, string cat, string msg, string? meta = null)
    {
        var entry = new LogEntry
        {
            Ts = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
            Level = level, Category = cat, Message = msg, Meta = meta
        };
        _recent.Enqueue(entry);
        while (_recent.Count > RecentCap && _recent.TryDequeue(out _)) { }
        try { OnEntry?.Invoke(entry); } catch { }
        if (!_queue.IsAddingCompleted) _queue.TryAdd(entry);
    }

    private void WorkerLoop()
    {
        var sb = new StringBuilder(4096);
        foreach (var entry in _queue.GetConsumingEnumerable(_cts.Token))
        {
            try
            {
                sb.Clear();
                var dt = DateTimeOffset.FromUnixTimeMilliseconds(entry.Ts).ToLocalTime();
                sb.Append(dt.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                  .Append(" [").Append(entry.Level.ToString().ToUpperInvariant()).Append("] ")
                  .Append(entry.Category).Append(": ").Append(entry.Message);
                if (entry.Meta is not null) sb.Append(" | ").Append(entry.Meta);
                sb.AppendLine();

                lock (_fileLock)
                {
                    if (new FileInfo(_currentFile).Exists && new FileInfo(_currentFile).Length > 5 * 1024 * 1024)
                        RollFile();
                    File.AppendAllText(_currentFile, sb.ToString(), Encoding.UTF8);
                }
            }
            catch { }
        }
    }

    private void RollFile()
    {
        var name = $"mailmonitor-{DateTime.Now:yyyyMMdd-HHmmss}.log";
        _currentFile = Path.Combine(_paths.LogsDir, name);
        try
        {
            var files = new DirectoryInfo(_paths.LogsDir).GetFiles("mailmonitor-*.log")
                .OrderByDescending(f => f.CreationTimeUtc).Skip(10).ToList();
            foreach (var f in files) { try { f.Delete(); } catch { } }
        }
        catch { }
    }

    public void Dispose()
    {
        try { _queue.CompleteAdding(); } catch { }
        try { _worker.Join(TimeSpan.FromSeconds(2)); } catch { }
        try { _cts.Cancel(); _cts.Dispose(); } catch { }
        _queue.Dispose();
    }
}
