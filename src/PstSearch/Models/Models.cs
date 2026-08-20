using System.Collections.Generic;

namespace PstSearchTool.Models
{
    /// <summary>Outlook 中的一個郵件存放區（PST/OST/Exchange）。</summary>
    public class StoreInfo
    {
        public string DisplayName;
        public string FilePath;
        public string StoreId;
        public bool IsDefault;
    }

    /// <summary>資料夾樹節點。</summary>
    public class FolderNode
    {
        public string Name;
        public string EntryId;
        public string Path;   // 完整路徑，以 '/' 分隔，例如：收件匣/客戶A
        public string Kind;   // inbox / sent / drafts / deleted / junk / outbox / other
        public List<FolderNode> Children = new List<FolderNode>();
    }

    /// <summary>要索引的資料夾（路徑 + 分類）。</summary>
    public class FolderInfo
    {
        public string Path;
        public string Kind;
    }

    /// <summary>從 Outlook 抽取出的單封郵件欄位。</summary>
    public class MailDoc
    {
        public string StoreId;
        public string StoreName;
        public string PstPath;
        public string FolderPath;
        public string FolderKind;
        public string Subject;
        public string FromName;
        public string FromEmail;
        public string ToList;
        public string CcList;
        public string ReceivedTime;   // yyyy-MM-dd HH:mm:ss
        public string EntryId;
        public string Body;
        public string SearchText;     // 供倒排索引使用的合併文字
    }

    /// <summary>搜尋結果列。</summary>
    public class SearchResultItem
    {
        public long Id;
        public string StoreName;
        public string FolderPath;
        public string Subject;
        public string FromName;
        public string FromEmail;
        public string ToList;
        public string ReceivedTime;
        public string EntryId;
        public string StoreId;
        public string Body;           // 已截斷（僅供摘要）
        public string Snippet;
    }
}
