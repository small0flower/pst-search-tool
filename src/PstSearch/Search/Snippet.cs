using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace PstSearchTool.Search
{
    /// <summary>從郵件文字中擷取含關鍵字的摘要片段。</summary>
    internal static class Snippet
    {
        public static string Make(string haystack, string[] keywords)
        {
            if (string.IsNullOrEmpty(haystack)) return "";
            string text = haystack.Replace("\r", " ").Replace("\n", " ");
            text = Regex.Replace(text, @"\s+", " ").Trim();
            int pos = -1;
            foreach (var kw in keywords.Where(k => !string.IsNullOrEmpty(k)))
            {
                int p = text.IndexOf(kw, StringComparison.Ordinal);
                if (p >= 0 && (pos < 0 || p < pos)) pos = p;
            }
            if (pos < 0)
                return text.Length > 200 ? text.Substring(0, 200) + "…" : text;
            int start = Math.Max(0, pos - 40);
            int len = Math.Min(text.Length - start, 200);
            return (start > 0 ? "…" : "") + text.Substring(start, len) + (start + len < text.Length ? "…" : "");
        }
    }
}
