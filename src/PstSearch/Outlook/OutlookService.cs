using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using PstSearchTool.Models;

namespace PstSearchTool.Outlook
{
    /// <summary>
    /// 以晚期繫結 (dynamic) 呼叫 Outlook COM，不依賴任何特定版本 PIA，
    /// 因此可相容 Outlook 2007 ~ 最新版。
    /// 位址類欄位一律透過 PropertyAccessor 讀取 MAPI 屬性，
    /// 避免觸發 Outlook 的「程式嘗試存取電子郵件地址資訊」安全性警告（OM Guard）。
    /// 注意：Outlook 2007 為 32 位元，本程式必須以 x86 建置才能 COM 互通。
    /// </summary>
    internal static class OutlookService
    {
        // OlDefaultFolders
        private const int OlFolderInbox = 6;
        private const int OlFolderSentMail = 5;
        private const int OlFolderDrafts = 16;
        private const int OlFolderDeletedItems = 3;
        private const int OlFolderOutbox = 4;
        private const int OlFolderJunk = 23;

        // MAPI proptag（PropertyAccessor 專用；讀取這些屬性不會觸發 OM Guard 警告）
        private static readonly string[] SenderNameTags =
        {
            "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F", // PR_SENDER_NAME_W
            "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E", // PR_SENDER_NAME_A
            "http://schemas.microsoft.com/mapi/proptag/0x0042001F", // PR_SENT_REPRESENTING_NAME_W
            "http://schemas.microsoft.com/mapi/proptag/0x0042001E", // PR_SENT_REPRESENTING_NAME_A
        };
        private static readonly string[] SenderEmailTags =
        {
            "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F", // PR_SENDER_EMAIL_ADDRESS_W
            "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E", // PR_SENDER_EMAIL_ADDRESS_A
            "http://schemas.microsoft.com/mapi/proptag/0x0065001F", // PR_SENT_REPRESENTING_EMAIL_ADDRESS_W
            "http://schemas.microsoft.com/mapi/proptag/0x0065001E", // PR_SENT_REPRESENTING_EMAIL_ADDRESS_A
        };
        private static readonly string[] ToTags =
        {
            "http://schemas.microsoft.com/mapi/proptag/0x0E04001F", // PR_DISPLAY_TO_W
            "http://schemas.microsoft.com/mapi/proptag/0x0E04001E", // PR_DISPLAY_TO_A
        };
        private static readonly string[] CcTags =
        {
            "http://schemas.microsoft.com/mapi/proptag/0x0E03001F", // PR_DISPLAY_CC_W
            "http://schemas.microsoft.com/mapi/proptag/0x0E03001E", // PR_DISPLAY_CC_A
        };

        private static readonly object LockObj = new object();

        /// <summary>取得 Outlook.Application（執行中則沿用，否則啟動）。</summary>
        public static dynamic GetApplication()
        {
            lock (LockObj)
            {
                try
                {
                    object running = Marshal.GetActiveObject("Outlook.Application");
                    return running;
                }
                catch
                {
                    Type t = Type.GetTypeFromProgID("Outlook.Application");
                    if (t == null)
                        throw new InvalidOperationException("找不到 Outlook.Application。請確認已安裝 Outlook（2007 或更新版本）。");
                    return Activator.CreateInstance(t);
                }
            }
        }

        /// <summary>列出現行 Outlook 設定檔內的所有郵件存放區。</summary>
        public static List<StoreInfo> GetStores()
        {
            var result = new List<StoreInfo>();
            dynamic app = null, ns = null, defaultStore = null, stores = null;
            try
            {
                app = GetApplication();
                ns = app.GetNamespace("MAPI");
                string defaultStoreId = "";
                try { defaultStore = ns.DefaultStore; defaultStoreId = ComUtil.SafeStr(defaultStore.StoreID); } catch { }
                stores = ns.Stores;
                int count = 0;
                try { count = Convert.ToInt32(stores.Count); } catch { }
                for (int i = 1; i <= count; i++)
                {
                    dynamic store = null;
                    try
                    {
                        store = stores[i];
                        result.Add(new StoreInfo
                        {
                            DisplayName = ComUtil.SafeStr(store.DisplayName),
                            FilePath = ComUtil.SafeStr(store.FilePath),
                            StoreId = ComUtil.SafeStr(store.StoreID),
                            IsDefault = ComUtil.SafeStr(store.StoreID) == defaultStoreId
                        });
                    }
                    catch { }
                    finally { ComUtil.Release(store); }
                }
                return result;
            }
            finally
            {
                ComUtil.Release(stores);
                ComUtil.Release(defaultStore);
                ComUtil.Release(ns);
                ComUtil.Release(app);
                ComUtil.CoCleanup();
            }
        }

        /// <summary>回傳指定 store 的資料夾樹（含分類 inbox/sent/...）。</summary>
        public static List<FolderNode> GetFolderTree(string storeId)
        {
            var result = new List<FolderNode>();
            dynamic app = null, ns = null, store = null, root = null;
            try
            {
                app = GetApplication();
                ns = app.GetNamespace("MAPI");
                store = FindStore(ns, storeId);
                if (store == null) throw new InvalidOperationException("找不到指定的郵件存放區。");

                // 預設資料夾 EntryID 對照（僅當此 store 是預設存放區時有效）
                var defaultIds = new Dictionary<string, string>();
                string defaultStoreId = "";
                try { defaultStoreId = ComUtil.SafeStr(ns.DefaultStore.StoreID); } catch { }
                if (ComUtil.SafeStr(store.StoreID) == defaultStoreId)
                {
                    try
                    {
                        defaultIds["inbox"] = ComUtil.SafeStr(ns.GetDefaultFolder(OlFolderInbox).EntryID);
                        defaultIds["sent"] = ComUtil.SafeStr(ns.GetDefaultFolder(OlFolderSentMail).EntryID);
                        defaultIds["drafts"] = ComUtil.SafeStr(ns.GetDefaultFolder(OlFolderDrafts).EntryID);
                        defaultIds["deleted"] = ComUtil.SafeStr(ns.GetDefaultFolder(OlFolderDeletedItems).EntryID);
                        defaultIds["outbox"] = ComUtil.SafeStr(ns.GetDefaultFolder(OlFolderOutbox).EntryID);
                        defaultIds["junk"] = ComUtil.SafeStr(ns.GetDefaultFolder(OlFolderJunk).EntryID);
                    }
                    catch { }
                }

                root = store.GetRootFolder();
                dynamic children = null;
                try
                {
                    children = root.Folders;
                    int n = Convert.ToInt32(children.Count);
                    for (int i = 1; i <= n; i++)
                    {
                        dynamic child = null;
                        try
                        {
                            child = children[i];
                            result.Add(BuildTree(child, "", defaultIds, 0));
                        }
                        catch { }
                        finally { ComUtil.Release(child); }
                    }
                }
                finally { ComUtil.Release(children); }
                return result;
            }
            finally
            {
                ComUtil.Release(root);
                ComUtil.Release(store);
                ComUtil.Release(ns);
                ComUtil.Release(app);
                ComUtil.CoCleanup();
            }
        }

        private static FolderNode BuildTree(dynamic folder, string parentPath, Dictionary<string, string> defaultIds, int depth)
        {
            var node = new FolderNode
            {
                Name = ComUtil.SafeStr(folder.Name),
                EntryId = ComUtil.SafeStr(folder.EntryID)
            };
            node.Path = string.IsNullOrEmpty(parentPath) ? node.Name : parentPath + "/" + node.Name;
            node.Kind = ClassifyFolder(node.EntryId, node.Name, defaultIds);
            if (depth >= 40) return node; // 深度保護
            dynamic children = null;
            try
            {
                children = folder.Folders;
                int n = Convert.ToInt32(children.Count);
                for (int i = 1; i <= n; i++)
                {
                    dynamic child = null;
                    try
                    {
                        child = children[i];
                        node.Children.Add(BuildTree(child, node.Path, defaultIds, depth + 1));
                    }
                    catch { }
                    finally { ComUtil.Release(child); }
                }
            }
            catch { }
            finally { ComUtil.Release(children); }
            return node;
        }

        private static string ClassifyFolder(string entryId, string name, Dictionary<string, string> defaultIds)
        {
            if (defaultIds != null)
            {
                foreach (var kv in defaultIds)
                    if (!string.IsNullOrEmpty(kv.Value) && kv.Value == entryId) return kv.Key;
            }
            string n = (name ?? "").Trim();
            if (n == "收件匣" || n == "收件夹" || n == "Inbox" || n == "INBOX") return "inbox";
            if (n == "寄件備份" || n == "寄件匣" || n == "寄件夹" || n == "Sent Items" || n == "Sent") return "sent";
            if (n == "草稿" || n == "Drafts") return "drafts";
            if (n == "刪除的郵件" || n == "删除的邮件" || n == "Deleted Items") return "deleted";
            if (n == "垃圾郵件" || n == "Junk E-mail" || n == "Junk") return "junk";
            if (n == "Outbox") return "outbox";
            return "other";
        }

        /// <summary>快速取得資料夾內郵件數（Items.Count，不做逐封讀取）。</summary>
        public static long CountFolderItems(string storeId, string folderPath)
        {
            dynamic app = null, ns = null, store = null, folder = null, items = null;
            try
            {
                app = GetApplication();
                ns = app.GetNamespace("MAPI");
                store = FindStore(ns, storeId);
                if (store == null) return 0;
                folder = FindFolderByPath(store, folderPath);
                if (folder == null) return 0;
                items = folder.Items;
                return Convert.ToInt64(items.Count);
            }
            catch { return 0; }
            finally
            {
                ComUtil.Release(items);
                ComUtil.Release(folder);
                ComUtil.Release(store);
                ComUtil.Release(ns);
                ComUtil.Release(app);
                ComUtil.CoCleanup();
            }
        }

        /// <summary>
        /// 逐封讀取資料夾內郵件並回呼 onItem（同步執行）。
        /// progress(done, totalInFolder) 每 25 封回呼一次。
        /// </summary>
        public static long ReadFolderItems(string storeId, string folderPath, string folderKind,
            Action<MailDoc> onItem, Func<bool> isCancelled, Action<long, long> progress)
        {
            dynamic app = null, ns = null, store = null, folder = null, items = null, item = null;
            long count = 0;
            try
            {
                app = GetApplication();
                ns = app.GetNamespace("MAPI");
                store = FindStore(ns, storeId);
                if (store == null) throw new InvalidOperationException("找不到郵件存放區：" + storeId);
                folder = FindFolderByPath(store, folderPath);
                if (folder == null) throw new InvalidOperationException("找不到資料夾：" + folderPath);
                items = folder.Items;
                long totalItems = 0;
                try { totalItems = Convert.ToInt64(items.Count); } catch { }
                item = items.GetFirst();
                while (item != null)
                {
                    if (isCancelled()) break;
                    MailDoc doc = ExtractItem(item, folderPath, folderKind);
                    if (doc != null) onItem(doc);
                    count++;
                    if (progress != null && (count % 25 == 0 || totalItems == 0 || count == totalItems))
                        progress(count, totalItems);
                    dynamic next = null;
                    try { next = items.GetNext(); } catch { next = null; }
                    ComUtil.Release(item);
                    item = next;
                }
                return count;
            }
            finally
            {
                ComUtil.Release(item);
                ComUtil.Release(items);
                ComUtil.Release(folder);
                ComUtil.Release(store);
                ComUtil.Release(ns);
                ComUtil.Release(app);
                ComUtil.CoCleanup();
            }
        }

        private static MailDoc ExtractItem(dynamic item, string folderPath, string folderKind)
        {
            try
            {
                var doc = new MailDoc
                {
                    FolderPath = folderPath,
                    FolderKind = folderKind,
                    Subject = ComUtil.SafeStr(item.Subject),
                    ReceivedTime = ComUtil.SafeDate(item.ReceivedTime),
                    EntryId = ComUtil.SafeStr(item.EntryID),
                    FromName = ReadProp(item, SenderNameTags),
                    FromEmail = ReadProp(item, SenderEmailTags),
                    ToList = ReadProp(item, ToTags),
                    CcList = ReadProp(item, CcTags),
                    Body = ComUtil.SafeStr(item.Body),
                    Attachments = ReadAttachmentNames(item)
                };
                if (string.IsNullOrEmpty(doc.ReceivedTime)) doc.ReceivedTime = ComUtil.SafeDate(item.SentOn);
                return doc;
            }
            catch { return null; }
        }

        /// <summary>讀取附件檔名（以 ; 分隔）。透過 Attachments 集合，不會觸發安全性警告。</summary>
        private static string ReadAttachmentNames(dynamic item)
        {
            try
            {
                dynamic ats = item.Attachments;
                if (ats == null) return "";
                int n = Convert.ToInt32(ats.Count);
                if (n <= 0) return "";
                var names = new List<string>();
                for (int i = 1; i <= n; i++)
                {
                    dynamic a = null;
                    try
                    {
                        a = ats[i];
                        string fn = ComUtil.SafeStr(a.FileName);
                        if (string.IsNullOrWhiteSpace(fn)) fn = ComUtil.SafeStr(a.DisplayName);
                        if (!string.IsNullOrWhiteSpace(fn)) names.Add(fn);
                    }
                    catch { }
                    finally { ComUtil.Release(a); }
                }
                return string.Join("; ", names);
            }
            catch { return ""; }
        }

        /// <summary>依序嘗試多個 MAPI proptag，回傳第一個非空值（避免 OM Guard 警告）。</summary>
        private static string ReadProp(dynamic item, string[] tags)
        {
            try
            {
                dynamic pa = item.PropertyAccessor;
                if (pa == null) return "";
                foreach (string tag in tags)
                {
                    try
                    {
                        object v = pa.GetProperty(tag);
                        if (v != null)
                        {
                            string s = Convert.ToString(v);
                            if (!string.IsNullOrWhiteSpace(s)) return s;
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return "";
        }

        /// <summary>將外部 PST 加入 Outlook 設定檔；已存在則回傳 false 不動作。</summary>
        public static bool AddStore(string pstPath)
        {
            dynamic app = null, ns = null, stores = null;
            try
            {
                app = GetApplication();
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                int n = Convert.ToInt32(stores.Count);
                for (int i = 1; i <= n; i++)
                {
                    dynamic store = null;
                    try
                    {
                        store = stores[i];
                        string fp = ComUtil.SafeStr(store.FilePath);
                        if (!string.IsNullOrEmpty(fp) &&
                            string.Equals(Path.GetFullPath(fp), Path.GetFullPath(pstPath), StringComparison.OrdinalIgnoreCase))
                            return false;
                    }
                    catch { }
                    finally { ComUtil.Release(store); }
                }
                ns.AddStore(pstPath);
                return true;
            }
            finally
            {
                ComUtil.Release(stores);
                ComUtil.Release(ns);
                ComUtil.Release(app);
                ComUtil.CoCleanup();
            }
        }

        /// <summary>以 EntryID 於 Outlook 開啟原件（該 PST 必須仍掛載於設定檔）。</summary>
        public static void OpenItem(string entryId, string storeId)
        {
            dynamic app = null, ns = null, item = null;
            try
            {
                app = GetApplication();
                ns = app.GetNamespace("MAPI");
                item = ns.GetItemFromID(entryId, storeId);
                if (item == null)
                    throw new InvalidOperationException("找不到該郵件。它所屬的 PST 可能已從 Outlook 移除，請先重新掛載（或重新建立索引）後再試。");
                item.Display(false);
            }
            finally
            {
                ComUtil.Release(item);
                ComUtil.Release(ns);
                ComUtil.Release(app);
                ComUtil.CoCleanup();
            }
        }

        private static dynamic FindStore(dynamic ns, string storeId)
        {
            dynamic stores = null;
            try
            {
                stores = ns.Stores;
                int n = Convert.ToInt32(stores.Count);
                for (int i = 1; i <= n; i++)
                {
                    dynamic store = null;
                    try
                    {
                        store = stores[i];
                        if (ComUtil.SafeStr(store.StoreID) == storeId) return store; // 呼叫者負責釋放
                        ComUtil.Release(store);
                    }
                    catch { ComUtil.Release(store); }
                }
                return null;
            }
            finally { ComUtil.Release(stores); }
        }

        private static dynamic FindFolderByPath(dynamic store, string folderPath)
        {
            string[] parts = folderPath.Split('/');
            dynamic current = null;
            try
            {
                current = store.GetRootFolder();
                foreach (string part in parts)
                {
                    dynamic children = null, next = null;
                    try
                    {
                        children = current.Folders;
                        int n = Convert.ToInt32(children.Count);
                        for (int i = 1; i <= n; i++)
                        {
                            dynamic child = children[i];
                            if (ComUtil.SafeStr(child.Name) == part)
                            {
                                next = child; // 保留給下一輪 / 回傳
                            }
                            else
                            {
                                ComUtil.Release(child);
                            }
                        }
                    }
                    finally { ComUtil.Release(children); }
                    if (next == null)
                    {
                        ComUtil.Release(current);
                        return null;
                    }
                    ComUtil.Release(current);
                    current = next;
                }
                return current; // 呼叫者負責釋放
            }
            catch
            {
                ComUtil.Release(current);
                return null;
            }
        }
    }
}
