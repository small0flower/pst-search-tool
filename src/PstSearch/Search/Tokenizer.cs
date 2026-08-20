using System;
using System.Collections.Generic;

namespace PstSearchTool.Search
{
    /// <summary>
    /// 自建倒排索引的分詞器（不依賴 SQLite FTS5，相容任何 SQLite 版本）：
    /// - CJK（非 ASCII 字母）連續段：產生 2 字元 bigram 與 3 字元 trigram；
    /// - ASCII 字母/數字連續段：小寫整詞（長度 ≥ 2）。
    /// 因此中文 2 字元以上的子字串都可精確搜尋（如「鴻海」），英文為不分大小寫的整詞比對。
    /// </summary>
    internal static class Tokenizer
    {
        private static bool IsCjk(char ch)
        {
            return char.IsLetter(ch) && ch > 127;
        }

        private static bool IsAsciiWordChar(char ch)
        {
            return ch < 128 && char.IsLetterOrDigit(ch);
        }

        /// <summary>索引用：產生一則訊息的所有 token（去重）。</summary>
        public static HashSet<string> IndexTokens(string text)
        {
            var toks = new HashSet<string>();
            if (string.IsNullOrEmpty(text)) return toks;
            int i = 0, n = text.Length;
            while (i < n)
            {
                char ch = text[i];
                if (IsCjk(ch))
                {
                    int j = i;
                    while (j < n && IsCjk(text[j])) j++;
                    string run = text.Substring(i, j - i);
                    int L = run.Length;
                    for (int k = 0; k < L - 1; k++) toks.Add(run.Substring(k, 2));
                    for (int k = 0; k < L - 2; k++) toks.Add(run.Substring(k, 3));
                    i = j;
                }
                else if (IsAsciiWordChar(ch))
                {
                    int j = i;
                    while (j < n && IsAsciiWordChar(text[j])) j++;
                    string w = text.Substring(i, j - i).ToLowerInvariant();
                    if (w.Length >= 2) toks.Add(w);
                    i = j;
                }
                else i++;
            }
            return toks;
        }

        /// <summary>
        /// 查詢用：產生關鍵字的查詢 token。
        /// fallback=true 表示含單字元（無法以索引精確比對），呼叫端應改用 instr 掃描。
        /// </summary>
        public static List<string> QueryTerms(string keyword, out bool fallback)
        {
            var terms = new List<string>();
            fallback = false;
            if (string.IsNullOrEmpty(keyword)) return terms;
            int i = 0, n = keyword.Length;
            while (i < n)
            {
                char ch = keyword[i];
                if (IsCjk(ch))
                {
                    int j = i;
                    while (j < n && IsCjk(keyword[j])) j++;
                    string run = keyword.Substring(i, j - i);
                    int L = run.Length;
                    if (L == 1) fallback = true;
                    else if (L == 2) terms.Add(run);
                    else
                    {
                        for (int k = 0; k < L - 2; k++) terms.Add(run.Substring(k, 3));
                    }
                    i = j;
                }
                else if (IsAsciiWordChar(ch))
                {
                    int j = i;
                    while (j < n && IsAsciiWordChar(keyword[j])) j++;
                    string w = keyword.Substring(i, j - i).ToLowerInvariant();
                    if (w.Length == 1) fallback = true;
                    else terms.Add(w);
                    i = j;
                }
                else i++;
            }
            return terms;
        }
    }
}
