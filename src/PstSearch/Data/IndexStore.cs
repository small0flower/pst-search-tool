using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using PstSearchTool.Models;
using PstSearchTool.Search;

namespace PstSearchTool.Data
{
    /// <summary>資料夾快照：用於增量判斷（郵件數 + PST 檔大小/修改時間）。</summary>
    internal class FolderSnapshot
    {
        public string StoreId;
        public string FolderPath;
        public long ItemCount;
        public long PstFileSize;
        public long PstFileMtime;
        public DateTime IndexedAt;
    }

    /// <summary>
    /// SQLite 索引儲存層。
    /// 全文搜尋採用自建倒排索引（postings 表：term → msg_id），
    /// 不依賴 SQLite FTS5 擴充（官方 System.Data.SQLite 二進位對 FTS5 的支援不一），
    /// 因此任何 SQLite 版本皆可正常運作；中文 2 字元以上子字串可精確搜尋。
    /// 所有操作以 lock 保護，單一連線、單一執行緒使用。
    /// </summary>
    internal class IndexStore : IDisposable
    {
        private const int PostingFlushThreshold = 50000;

        private readonly SQLiteConnection _conn;
        private readonly object _sync = new object();
        private SQLiteTransaction _tx;
        private readonly List<KeyValuePair<string, long>> _pendingPostings = new List<KeyValuePair<string, long>>();

        public IndexStore(string dbPath)
        {
            string dir = Path.GetDirectoryName(dbPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            var b = new SQLiteConnectionStringBuilder
            {
                DataSource = dbPath,
                Version = 3,
                JournalMode = SQLiteJournalModeEnum.Wal,
                Pooling = false
            };
            _conn = new SQLiteConnection(b.ToString());
            _conn.Open();
            EnsureSchema();
        }

        public void Dispose()
        {
            lock (_sync)
            {
                _pendingPostings.Clear();
                try { if (_tx != null) { _tx.Dispose(); _tx = null; } } catch { }
                try { _conn.Close(); } catch { }
                try { _conn.Dispose(); } catch { }
            }
        }

        private void EnsureSchema()
        {
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS messages(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  store_id TEXT NOT NULL,
  store_name TEXT,
  pst_path TEXT,
  folder_path TEXT NOT NULL,
  folder_kind TEXT,
  subject TEXT,
  from_name TEXT,
  from_email TEXT,
  to_list TEXT,
  cc_list TEXT,
  received_time TEXT,
  entry_id TEXT NOT NULL,
  body TEXT,
  attachments TEXT,
  search_text TEXT,
  indexed_at TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_msg ON messages(store_id, entry_id);
CREATE INDEX IF NOT EXISTS ix_msg_folder ON messages(folder_path);
CREATE INDEX IF NOT EXISTS ix_msg_time ON messages(received_time);
CREATE INDEX IF NOT EXISTS ix_msg_from ON messages(from_email);
CREATE TABLE IF NOT EXISTS postings(
  term TEXT NOT NULL,
  msg_id INTEGER NOT NULL,
  PRIMARY KEY(term, msg_id)
) WITHOUT ROWID;
CREATE INDEX IF NOT EXISTS ix_postings_msg ON postings(msg_id);
CREATE TABLE IF NOT EXISTS folder_snapshots(
  store_id TEXT NOT NULL,
  folder_path TEXT NOT NULL,
  item_count INTEGER NOT NULL DEFAULT 0,
  pst_file_size INTEGER NOT NULL DEFAULT 0,
  pst_file_mtime INTEGER NOT NULL DEFAULT 0,
  indexed_at TEXT,
  PRIMARY KEY(store_id, folder_path)
);";
                cmd.ExecuteNonQuery();
            }
            // 遷移：為既有資料庫補上 attachments 欄位（V2 附件檔名索引）
            try
            {
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = "ALTER TABLE messages ADD COLUMN attachments TEXT";
                    cmd.ExecuteNonQuery();
                }
            }
            catch { /* 欄位已存在則忽略 */ }
        }

        public void BeginTransaction()
        {
            lock (_sync)
            {
                if (_tx != null) return;
                _tx = _conn.BeginTransaction();
            }
        }

        public void Commit()
        {
            lock (_sync)
            {
                if (_tx == null) return;
                try
                {
                    FlushPostings();
                    _tx.Commit();
                }
                finally { _tx.Dispose(); _tx = null; }
            }
        }

        public void Rollback()
        {
            lock (_sync)
            {
                _pendingPostings.Clear();
                if (_tx == null) return;
                try { _tx.Rollback(); }
                finally { _tx.Dispose(); _tx = null; }
            }
        }

        public void InsertMessage(MailDoc doc)
        {
            lock (_sync)
            {
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.Transaction = _tx;
                    cmd.CommandText = @"
INSERT INTO messages(store_id, store_name, pst_path, folder_path, folder_kind, subject, from_name, from_email, to_list, cc_list, received_time, entry_id, body, attachments, search_text, indexed_at)
VALUES(@s, @sn, @p, @f, @k, @sub, @fn, @fe, @to, @cc, @t, @e, @b, @at, @st, @ia)
ON CONFLICT(store_id, entry_id) DO UPDATE SET
 store_name=@sn, pst_path=@p, folder_path=@f, folder_kind=@k, subject=@sub,
 from_name=@fn, from_email=@fe, to_list=@to, cc_list=@cc, received_time=@t,
 body=@b, attachments=@at, search_text=@st, indexed_at=@ia";
                    cmd.Parameters.AddWithValue("@s", doc.StoreId ?? "");
                    cmd.Parameters.AddWithValue("@sn", doc.StoreName ?? "");
                    cmd.Parameters.AddWithValue("@p", doc.PstPath ?? "");
                    cmd.Parameters.AddWithValue("@f", doc.FolderPath ?? "");
                    cmd.Parameters.AddWithValue("@k", doc.FolderKind ?? "other");
                    cmd.Parameters.AddWithValue("@sub", doc.Subject ?? "");
                    cmd.Parameters.AddWithValue("@fn", doc.FromName ?? "");
                    cmd.Parameters.AddWithValue("@fe", doc.FromEmail ?? "");
                    cmd.Parameters.AddWithValue("@to", doc.ToList ?? "");
                    cmd.Parameters.AddWithValue("@cc", doc.CcList ?? "");
                    cmd.Parameters.AddWithValue("@t", doc.ReceivedTime ?? "");
                    cmd.Parameters.AddWithValue("@e", doc.EntryId ?? "");
                    cmd.Parameters.AddWithValue("@b", doc.Body ?? "");
                    cmd.Parameters.AddWithValue("@at", doc.Attachments ?? "");
                    cmd.Parameters.AddWithValue("@st", doc.SearchText ?? "");
                    cmd.Parameters.AddWithValue("@ia", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.ExecuteNonQuery();
                }
                // 反查 id（UPSERT 更新後 last_insert_rowid() 會回傳舊值，不可信）
                long id = 0;
                using (var q = _conn.CreateCommand())
                {
                    q.Transaction = _tx;
                    q.CommandText = "SELECT id FROM messages WHERE store_id=@s AND entry_id=@e";
                    q.Parameters.AddWithValue("@s", doc.StoreId ?? "");
                    q.Parameters.AddWithValue("@e", doc.EntryId ?? "");
                    object v = q.ExecuteScalar();
                    if (v != null) id = Convert.ToInt64(v);
                }
                if (id > 0)
                {
                    foreach (string t in Tokenizer.IndexTokens(doc.SearchText ?? ""))
                        _pendingPostings.Add(new KeyValuePair<string, long>(t, id));
                    FlushPostingsIfNeeded();
                }
            }
        }

        private void FlushPostingsIfNeeded()
        {
            if (_pendingPostings.Count >= PostingFlushThreshold) FlushPostings();
        }

        private void FlushPostings()
        {
            if (_pendingPostings.Count == 0) return;
            using (var cmd = _conn.CreateCommand())
            {
                cmd.Transaction = _tx;
                int idx = 0;
                while (idx < _pendingPostings.Count)
                {
                    int take = Math.Min(200, _pendingPostings.Count - idx);
                    var sb = new StringBuilder("INSERT INTO postings(term, msg_id) VALUES ");
                    for (int k = 0; k < take; k++)
                    {
                        if (k > 0) sb.Append(",");
                        sb.Append("(@t").Append(k).Append(",@m").Append(k).Append(")");
                    }
                    cmd.CommandText = sb.ToString();
                    cmd.Parameters.Clear();
                    for (int k = 0; k < take; k++)
                    {
                        cmd.Parameters.AddWithValue("@t" + k, _pendingPostings[idx + k].Key);
                        cmd.Parameters.AddWithValue("@m" + k, _pendingPostings[idx + k].Value);
                    }
                    cmd.ExecuteNonQuery();
                    idx += take;
                }
            }
            _pendingPostings.Clear();
        }

        /// <summary>刪除某 store 下指定資料夾（含子資料夾）的索引與 postings。</summary>
        public void DeleteFolder(string storeId, string folderPath)
        {
            lock (_sync)
            {
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.Transaction = _tx;
                    cmd.CommandText = @"DELETE FROM postings WHERE msg_id IN
(SELECT id FROM messages WHERE store_id=@s AND (folder_path=@f OR folder_path LIKE @prefix))";
                    cmd.Parameters.AddWithValue("@s", storeId);
                    cmd.Parameters.AddWithValue("@f", folderPath);
                    cmd.Parameters.AddWithValue("@prefix", folderPath + "/%");
                    cmd.ExecuteNonQuery();
                }
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.Transaction = _tx;
                    cmd.CommandText = "DELETE FROM messages WHERE store_id=@s AND (folder_path=@f OR folder_path LIKE @prefix)";
                    cmd.Parameters.AddWithValue("@s", storeId);
                    cmd.Parameters.AddWithValue("@f", folderPath);
                    cmd.Parameters.AddWithValue("@prefix", folderPath + "/%");
                    cmd.ExecuteNonQuery();
                }
                // 同步清除該資料夾的快照（避免殘留導致清除邏輯重複判斷）
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.Transaction = _tx;
                    cmd.CommandText = "DELETE FROM folder_snapshots WHERE store_id=@s AND (folder_path=@f OR folder_path LIKE @prefix)";
                    cmd.Parameters.AddWithValue("@s", storeId);
                    cmd.Parameters.AddWithValue("@f", folderPath);
                    cmd.Parameters.AddWithValue("@prefix", folderPath + "/%");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public FolderSnapshot GetSnapshot(string storeId, string folderPath)
        {
            lock (_sync)
            {
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT item_count, pst_file_size, pst_file_mtime, indexed_at FROM folder_snapshots WHERE store_id=@s AND folder_path=@f";
                    cmd.Parameters.AddWithValue("@s", storeId);
                    cmd.Parameters.AddWithValue("@f", folderPath);
                    using (var r = cmd.ExecuteReader())
                    {
                        if (!r.Read()) return null;
                        var snap = new FolderSnapshot
                        {
                            StoreId = storeId,
                            FolderPath = folderPath,
                            ItemCount = r.GetInt64(0),
                            PstFileSize = r.GetInt64(1),
                            PstFileMtime = r.GetInt64(2)
                        };
                        DateTime d;
                        if (DateTime.TryParse(r.GetString(3), out d)) snap.IndexedAt = d;
                        return snap;
                    }
                }
            }
        }

        public void SaveSnapshot(FolderSnapshot s)
        {
            lock (_sync)
            {
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO folder_snapshots(store_id, folder_path, item_count, pst_file_size, pst_file_mtime, indexed_at)
VALUES(@s, @f, @c, @sz, @mt, @t)
ON CONFLICT(store_id, folder_path) DO UPDATE SET
 item_count=@c, pst_file_size=@sz, pst_file_mtime=@mt, indexed_at=@t";
                    cmd.Parameters.AddWithValue("@s", s.StoreId);
                    cmd.Parameters.AddWithValue("@f", s.FolderPath);
                    cmd.Parameters.AddWithValue("@c", s.ItemCount);
                    cmd.Parameters.AddWithValue("@sz", s.PstFileSize);
                    cmd.Parameters.AddWithValue("@mt", s.PstFileMtime);
                    cmd.Parameters.AddWithValue("@t", s.IndexedAt.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public long CountMessages()
        {
            lock (_sync)
            {
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM messages";
                    return Convert.ToInt64(cmd.ExecuteScalar());
                }
            }
        }

        /// <summary>回傳某存放區目前已索引的資料夾（依 folder_snapshots）。</summary>
        public List<string> GetIndexedFolders(string storeId)
        {
            lock (_sync)
            {
                var list = new List<string>();
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT folder_path FROM folder_snapshots WHERE store_id=@s";
                    cmd.Parameters.AddWithValue("@s", storeId);
                    using (var r = cmd.ExecuteReader())
                        while (r.Read()) list.Add(r.GetString(0));
                }
                return list;
            }
        }

        /// <summary>統計：寄件者 Top20、依資料夾、依月份。</summary>
        public StatsResult GetStats()
        {
            lock (_sync)
            {
                var res = new StatsResult();
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM messages";
                    res.Total = Convert.ToInt64(cmd.ExecuteScalar());
                }
                Action<string, List<KeyValuePair<string, long>>> query = (sql, target) =>
                {
                    using (var cmd = _conn.CreateCommand())
                    {
                        cmd.CommandText = sql;
                        using (var r = cmd.ExecuteReader())
                            while (r.Read())
                                target.Add(new KeyValuePair<string, long>(r.GetString(0), r.GetInt64(1)));
                    }
                };
                if (res.Total > 0)
                {
                    query("SELECT from_email, COUNT(*) FROM messages WHERE from_email <> '' GROUP BY from_email ORDER BY COUNT(*) DESC LIMIT 20", res.TopSenders);
                    query("SELECT folder_path, COUNT(*) FROM messages GROUP BY folder_path ORDER BY COUNT(*) DESC", res.FolderCounts);
                    query("SELECT substr(received_time,1,7), COUNT(*) FROM messages WHERE received_time <> '' GROUP BY substr(received_time,1,7) ORDER BY substr(received_time,1,7)", res.MonthCounts);
                }
                return res;
            }
        }

        /// <summary>
        /// 全文搜尋。
        /// 關鍵字：CJK 2 字元走 bigram、3 字元以上走 trigram（精確子字串）；英文整詞不分大小寫；
        /// 含單字元關鍵字時全部改為 instr 子字串掃描（較慢但正確）。
        /// folderPaths：null 表示不限；folderKind：inbox/sent/... 過濾；兩者可同時指定。
        /// </summary>
        public List<SearchResultItem> Search(string[] keywords, List<string> folderPaths,
            DateTime? from, DateTime? to, string senderEmail, string folderKind, int limit)
        {
            lock (_sync)
            {
                var sb = new StringBuilder();
                sb.Append(@"SELECT m.id, m.store_name, m.folder_path, m.subject, m.from_name, m.from_email, m.to_list, m.received_time, m.entry_id, m.store_id, substr(m.body, 1, 20000), m.attachments
FROM messages m ");

                var pars = new List<SQLiteParameter>();
                bool hasFilter = false;

                if (folderPaths != null && folderPaths.Count > 0)
                {
                    var parts = new List<string>();
                    int i = 0;
                    foreach (var fp in folderPaths)
                    {
                        parts.Add("(m.folder_path=@fp" + i + " OR m.folder_path LIKE @fpp" + i + ")");
                        pars.Add(new SQLiteParameter("@fp" + i, fp));
                        pars.Add(new SQLiteParameter("@fpp" + i, fp + "/%"));
                        i++;
                    }
                    sb.Append("WHERE (" + string.Join(" OR ", parts) + ") ");
                    hasFilter = true;
                }

                if (!string.IsNullOrEmpty(folderKind))
                {
                    sb.Append((hasFilter ? "AND " : "WHERE ") + "m.folder_kind=@kind ");
                    pars.Add(new SQLiteParameter("@kind", folderKind));
                    hasFilter = true;
                }

                if (from.HasValue)
                {
                    sb.Append((hasFilter ? "AND " : "WHERE ") + "m.received_time >= @from ");
                    pars.Add(new SQLiteParameter("@from", from.Value.ToString("yyyy-MM-dd")));
                    hasFilter = true;
                }
                if (to.HasValue)
                {
                    sb.Append((hasFilter ? "AND " : "WHERE ") + "m.received_time < @to ");
                    pars.Add(new SQLiteParameter("@to", to.Value.AddDays(1).ToString("yyyy-MM-dd")));
                    hasFilter = true;
                }
                if (!string.IsNullOrWhiteSpace(senderEmail))
                {
                    sb.Append((hasFilter ? "AND " : "WHERE ") + "(m.from_email LIKE @se OR m.from_name LIKE @se2) ");
                    pars.Add(new SQLiteParameter("@se", "%" + senderEmail + "%"));
                    pars.Add(new SQLiteParameter("@se2", "%" + senderEmail + "%"));
                    hasFilter = true;
                }

                string[] kws = (keywords ?? new string[0]).Where(k => !string.IsNullOrWhiteSpace(k)).ToArray();
                if (kws.Length > 0)
                {
                    var terms = new List<string>();
                    bool fallback = false;
                    foreach (var kw in kws)
                    {
                        terms.AddRange(Tokenizer.QueryTerms(kw, out bool fb));
                        fallback = fallback || fb;
                    }
                    terms = terms.Distinct().ToList();
                    if (!fallback && terms.Count > 0)
                    {
                        // 倒排索引路徑：所有查詢 token 必須同時出現
                        var ph = new List<string>();
                        for (int i = 0; i < terms.Count; i++)
                        {
                            ph.Add("@term" + i);
                            pars.Add(new SQLiteParameter("@term" + i, terms[i]));
                        }
                        sb.Append((hasFilter ? "AND " : "WHERE ") +
                            "m.id IN (SELECT p.msg_id FROM postings p WHERE p.term IN (" +
                            string.Join(",", ph) +
                            ") GROUP BY p.msg_id HAVING COUNT(DISTINCT p.term)=" + terms.Count + ") ");
                    }
                    else
                    {
                        // 含單字元關鍵字 → instr 全欄位掃描
                        var kparts = new List<string>();
                        int k = 0;
                        foreach (var kw in kws)
                        {
                            kparts.Add("(instr(m.subject, @kw" + k + ") > 0 OR instr(m.body, @kw" + k + "b) > 0 OR instr(m.from_name, @kw" + k + "n) > 0 OR instr(m.from_email, @kw" + k + "e) > 0 OR instr(m.to_list, @kw" + k + "t) > 0)");
                            pars.Add(new SQLiteParameter("@kw" + k, kw));
                            pars.Add(new SQLiteParameter("@kw" + k + "b", kw));
                            pars.Add(new SQLiteParameter("@kw" + k + "n", kw));
                            pars.Add(new SQLiteParameter("@kw" + k + "e", kw));
                            pars.Add(new SQLiteParameter("@kw" + k + "t", kw));
                            k++;
                        }
                        sb.Append((hasFilter ? "AND " : "WHERE ") + "(" + string.Join(" AND ", kparts) + ") ");
                    }
                }

                sb.Append("ORDER BY m.received_time DESC LIMIT @lim ");
                pars.Add(new SQLiteParameter("@lim", limit));

                var results = new List<SearchResultItem>();
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = sb.ToString();
                    cmd.Parameters.AddRange(pars.ToArray());
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            results.Add(new SearchResultItem
                            {
                                Id = r.GetInt64(0),
                                StoreName = r.IsDBNull(1) ? "" : r.GetString(1),
                                FolderPath = r.IsDBNull(2) ? "" : r.GetString(2),
                                Subject = r.IsDBNull(3) ? "" : r.GetString(3),
                                FromName = r.IsDBNull(4) ? "" : r.GetString(4),
                                FromEmail = r.IsDBNull(5) ? "" : r.GetString(5),
                                ToList = r.IsDBNull(6) ? "" : r.GetString(6),
                                ReceivedTime = r.IsDBNull(7) ? "" : r.GetString(7),
                                EntryId = r.IsDBNull(8) ? "" : r.GetString(8),
                                StoreId = r.IsDBNull(9) ? "" : r.GetString(9),
                                Body = r.IsDBNull(10) ? "" : r.GetString(10),
                                Attachments = r.IsDBNull(11) ? "" : r.GetString(11)
                            });
                        }
                    }
                }
                foreach (var it in results)
                    it.Snippet = Snippet.Make(it.Subject + "\n" + it.Body, kws);
                return results;
            }
        }
    }
}
