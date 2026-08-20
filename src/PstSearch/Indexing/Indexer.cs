using System;
using System.Collections.Generic;
using System.IO;
using PstSearchTool.Data;
using PstSearchTool.Models;
using PstSearchTool.Outlook;

namespace PstSearchTool.Indexing
{
    /// <summary>
    /// 協調 Outlook 讀取與 SQLite 寫入。
    /// 增量策略：若資料夾的郵件數與上次相同、且 PST 檔大小/修改時間未變 → 略過；
    /// 否則完整重建該資料夾索引（含其子資料夾路徑由呼叫端展開）。
    /// 取消時：已完成資料夾已提交，目前資料夾的未完成部分回滾，下次重跑即可補齊。
    /// </summary>
    internal class Indexer
    {
        private readonly IndexStore _db;

        public Indexer(IndexStore db) { _db = db; }

        public event Action<string, string> StatusChanged;        // (folderPath, message)
        public event Action<string, long, long> FolderProgress;   // (folderPath, done, totalInFolder)

        public long Run(StoreInfo store, List<FolderInfo> folders, bool forceRebuild, Func<bool> isCancelled)
        {
            long total = 0;
            foreach (var fi in folders)
            {
                if (isCancelled()) break;
                StatusChanged?.Invoke(fi.Path, "開始");
                FolderSnapshot snap = _db.GetSnapshot(store.StoreId, fi.Path);
                if (!forceRebuild && snap != null)
                {
                    long currentCount = OutlookService.CountFolderItems(store.StoreId, fi.Path);
                    if (ShouldSkip(store, snap, currentCount))
                    {
                        StatusChanged?.Invoke(fi.Path, "未變更，略過");
                        continue;
                    }
                }
                _db.DeleteFolder(store.StoreId, fi.Path);
                _db.BeginTransaction();
                long folderTotal = 0;
                bool cancelled = false;
                try
                {
                    OutlookService.ReadFolderItems(store.StoreId, fi.Path, fi.Kind,
                        doc =>
                        {
                            if (doc == null) return;
                            doc.StoreId = store.StoreId;
                            doc.StoreName = store.DisplayName;
                            doc.PstPath = store.FilePath;
                            doc.SearchText = BuildSearchText(doc);
                            _db.InsertMessage(doc);
                            folderTotal++;
                            total++;
                        },
                        () => isCancelled(),
                        (done, all) => FolderProgress?.Invoke(fi.Path, done, all));
                    cancelled = isCancelled();
                    if (!cancelled)
                    {
                        _db.Commit();
                        _db.SaveSnapshot(new FolderSnapshot
                        {
                            StoreId = store.StoreId,
                            FolderPath = fi.Path,
                            ItemCount = folderTotal,
                            PstFileSize = FileSize(store.FilePath),
                            PstFileMtime = FileMtime(store.FilePath),
                            IndexedAt = DateTime.Now
                        });
                        StatusChanged?.Invoke(fi.Path, "完成：" + folderTotal + " 封");
                    }
                    else
                    {
                        _db.Rollback();
                        StatusChanged?.Invoke(fi.Path, "已取消（此資料夾未儲存，下次重跑）");
                    }
                }
                catch
                {
                    try { _db.Rollback(); } catch { }
                    throw;
                }
            }
            return total;
        }

        private static bool ShouldSkip(StoreInfo store, FolderSnapshot snap, long currentCount)
        {
            if (snap == null) return false;
            if (snap.ItemCount != currentCount) return false;
            try
            {
                if (!string.IsNullOrEmpty(store.FilePath))
                {
                    var fi = new FileInfo(store.FilePath);
                    if (fi.Exists)
                    {
                        if (snap.PstFileSize != fi.Length) return false;
                        if (snap.PstFileMtime != fi.LastWriteTimeUtc.Ticks) return false;
                    }
                }
            }
            catch { return false; }
            return true;
        }

        private static string BuildSearchText(MailDoc doc)
        {
            return string.Join("\n",
                doc.Subject ?? "", doc.FromName ?? "", doc.FromEmail ?? "",
                doc.ToList ?? "", doc.CcList ?? "", doc.Body ?? "");
        }

        private static long FileSize(string path)
        {
            try { return string.IsNullOrEmpty(path) ? 0 : new FileInfo(path).Length; }
            catch { return 0; }
        }

        private static long FileMtime(string path)
        {
            try { return string.IsNullOrEmpty(path) ? 0 : new FileInfo(path).LastWriteTimeUtc.Ticks; }
            catch { return 0; }
        }
    }
}
