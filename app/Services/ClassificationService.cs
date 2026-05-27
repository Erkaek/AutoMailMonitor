using System.Globalization;
using Microsoft.Data.Sqlite;

namespace MailMonitor.Services;

public sealed class ClassificationService
{
    private readonly StorageService _storage;
    private readonly LogService _log;

    public const string CatDeclarations = "declarations";
    public const string CatReglements   = "reglements";
    public const string CatMails        = "mails";

    public ClassificationService(StorageService storage, LogService log)
    {
        _storage = storage; _log = log;
    }

    public static string ClassifyByPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return CatMails;
        var p = path.ToLowerInvariant();
        if (p.Contains("déclaration") || p.Contains("declaration") || p.Contains("décla") || p.Contains("decla"))
            return CatDeclarations;
        if (p.Contains("règlement") || p.Contains("reglement") || p.Contains("règl") || p.Contains("regl"))
            return CatReglements;
        return CatMails;
    }

    public static (int year, int week) IsoWeek(DateTime dt)
    {
        var cal = CultureInfo.InvariantCulture.Calendar;
        var day = cal.GetDayOfWeek(dt);
        if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday) dt = dt.AddDays(3);
        var week = cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        return (dt.Year, week);
    }

    public void UpsertEmail(long folderId, string folderCategory, OutlookMailItem mail)
    {
        var (y, w) = IsoWeek(mail.ReceivedTime);
        var ts = new DateTimeOffset(mail.ReceivedTime).ToUnixTimeSeconds();
        lock (_storage.WriteLock)
        {
            using var cmd = _storage.Connection.CreateCommand();
            cmd.CommandText = @"
INSERT INTO emails(folder_id, entry_id, subject, sender, received_ts, is_unread, category, iso_year, iso_week)
VALUES($fid, $eid, $sub, $snd, $rts, $unr, $cat, $y, $w)
ON CONFLICT(entry_id) DO UPDATE SET
  subject=excluded.subject, sender=excluded.sender, received_ts=excluded.received_ts,
  is_unread=excluded.is_unread, category=excluded.category, iso_year=excluded.iso_year, iso_week=excluded.iso_week";
            cmd.Parameters.AddWithValue("$fid", folderId);
            cmd.Parameters.AddWithValue("$eid", mail.EntryId);
            cmd.Parameters.AddWithValue("$sub", (object?)mail.Subject ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$snd", (object?)mail.Sender ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$rts", ts);
            cmd.Parameters.AddWithValue("$unr", mail.IsUnread ? 1 : 0);
            cmd.Parameters.AddWithValue("$cat", folderCategory);
            cmd.Parameters.AddWithValue("$y", y);
            cmd.Parameters.AddWithValue("$w", w);
            cmd.ExecuteNonQuery();
        }
    }

    public void BulkUpsertEmails(long folderId, string folderCategory, IReadOnlyList<OutlookMailItem> mails)
    {
        if (mails.Count == 0) return;
        lock (_storage.WriteLock)
        {
            using var tx = _storage.Connection.BeginTransaction();
            using var cmd = _storage.Connection.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = @"
INSERT INTO emails(folder_id, entry_id, subject, sender, received_ts, is_unread, category, iso_year, iso_week)
VALUES($fid, $eid, $sub, $snd, $rts, $unr, $cat, $y, $w)
ON CONFLICT(entry_id) DO UPDATE SET
  subject=excluded.subject, sender=excluded.sender, received_ts=excluded.received_ts,
  is_unread=excluded.is_unread, category=excluded.category, iso_year=excluded.iso_year, iso_week=excluded.iso_week";

            var pFid = cmd.Parameters.Add("$fid", SqliteType.Integer);
            var pEid = cmd.Parameters.Add("$eid", SqliteType.Text);
            var pSub = cmd.Parameters.Add("$sub", SqliteType.Text);
            var pSnd = cmd.Parameters.Add("$snd", SqliteType.Text);
            var pRts = cmd.Parameters.Add("$rts", SqliteType.Integer);
            var pUnr = cmd.Parameters.Add("$unr", SqliteType.Integer);
            var pCat = cmd.Parameters.Add("$cat", SqliteType.Text);
            var pY = cmd.Parameters.Add("$y", SqliteType.Integer);
            var pW = cmd.Parameters.Add("$w", SqliteType.Integer);

            foreach (var m in mails)
            {
                var (y, w) = IsoWeek(m.ReceivedTime);
                pFid.Value = folderId;
                pEid.Value = m.EntryId;
                pSub.Value = (object?)m.Subject ?? DBNull.Value;
                pSnd.Value = (object?)m.Sender ?? DBNull.Value;
                pRts.Value = new DateTimeOffset(m.ReceivedTime).ToUnixTimeSeconds();
                pUnr.Value = m.IsUnread ? 1 : 0;
                pCat.Value = folderCategory;
                pY.Value = y; pW.Value = w;
                cmd.ExecuteNonQuery();
            }
            tx.Commit();
        }
    }
}
