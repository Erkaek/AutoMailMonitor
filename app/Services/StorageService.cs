using Microsoft.Data.Sqlite;

namespace MailMonitor.Services;

public sealed class StorageService : IDisposable
{
    private readonly AppPaths _paths;
    private readonly LogService _log;
    private SqliteConnection _conn = default!;
    private readonly object _writeLock = new();

    public StorageService(AppPaths paths, LogService log) { _paths = paths; _log = log; }

    public void Initialize()
    {
        var cs = new SqliteConnectionStringBuilder
        {
            DataSource = _paths.DatabasePath,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Cache = SqliteCacheMode.Shared,
            Pooling = false
        }.ToString();

        _conn = new SqliteConnection(cs);
        _conn.Open();

        Exec("PRAGMA journal_mode=WAL;");
        Exec("PRAGMA synchronous=NORMAL;");
        Exec("PRAGMA temp_store=MEMORY;");
        Exec("PRAGMA mmap_size=268435456;");
        Exec("PRAGMA cache_size=-65536;");
        Exec("PRAGMA foreign_keys=ON;");

        CreateSchema();
        _log.Info("DB", "SQLite initialisée: " + _paths.DatabasePath);
    }

    private void CreateSchema()
    {
        Exec(@"
CREATE TABLE IF NOT EXISTS folders (
  id INTEGER PRIMARY KEY,
  store_id      TEXT NOT NULL,
  entry_id      TEXT NOT NULL,
  path          TEXT NOT NULL,
  display_name  TEXT,
  category      TEXT NOT NULL DEFAULT 'mails',
  is_monitored  INTEGER NOT NULL DEFAULT 1,
  last_scan_ts  INTEGER,
  last_received_ts INTEGER,
  UNIQUE(store_id, entry_id)
);
CREATE INDEX IF NOT EXISTS idx_folders_monitored ON folders(is_monitored);

CREATE TABLE IF NOT EXISTS emails (
  id INTEGER PRIMARY KEY,
  folder_id    INTEGER NOT NULL,
  entry_id     TEXT NOT NULL UNIQUE,
  subject      TEXT,
  sender       TEXT,
  received_ts  INTEGER NOT NULL,
  is_unread    INTEGER NOT NULL DEFAULT 0,
  category     TEXT,
  iso_year     INTEGER,
  iso_week     INTEGER,
  FOREIGN KEY(folder_id) REFERENCES folders(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_emails_folder ON emails(folder_id);
CREATE INDEX IF NOT EXISTS idx_emails_received ON emails(received_ts DESC);
CREATE INDEX IF NOT EXISTS idx_emails_week ON emails(iso_year, iso_week);
CREATE INDEX IF NOT EXISTS idx_emails_category ON emails(category);

CREATE TABLE IF NOT EXISTS weekly_comments (
  id INTEGER PRIMARY KEY,
  iso_year     INTEGER NOT NULL,
  iso_week     INTEGER NOT NULL,
  category     TEXT,
  comment_text TEXT NOT NULL,
  created_ts   INTEGER NOT NULL,
  updated_ts   INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_comments_week ON weekly_comments(iso_year, iso_week);

CREATE TABLE IF NOT EXISTS settings (
  key   TEXT PRIMARY KEY,
  value TEXT
);
");
    }

    private void Exec(string sql)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();
    }

    public SqliteConnection Connection => _conn;
    public object WriteLock => _writeLock;

    public string? GetSetting(string key)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = "SELECT value FROM settings WHERE key=$k";
        cmd.Parameters.AddWithValue("$k", key);
        return cmd.ExecuteScalar() as string;
    }

    public void SetSetting(string key, string value)
    {
        lock (_writeLock)
        {
            using var cmd = _conn.CreateCommand();
            cmd.CommandText = "INSERT INTO settings(key,value) VALUES($k,$v) ON CONFLICT(key) DO UPDATE SET value=excluded.value";
            cmd.Parameters.AddWithValue("$k", key);
            cmd.Parameters.AddWithValue("$v", value);
            cmd.ExecuteNonQuery();
        }
    }

    public void Dispose()
    {
        try { _conn?.Close(); _conn?.Dispose(); } catch { }
    }
}
