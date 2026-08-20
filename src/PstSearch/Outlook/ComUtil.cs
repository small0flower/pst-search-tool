using System;
using System.Runtime.InteropServices;

namespace PstSearchTool.Outlook
{
    /// <summary>COM 輔助：安全取值與 RCW 釋放。</summary>
    internal static class ComUtil
    {
        public static string SafeStr(dynamic value)
        {
            try
            {
                if (value == null) return "";
                object o = (object)value;
                return o == null ? "" : Convert.ToString(o) ?? "";
            }
            catch { return ""; }
        }

        public static string SafeDate(dynamic value)
        {
            try
            {
                if (value == null) return "";
                var d = (DateTime)value;
                return d.ToString("yyyy-MM-dd HH:mm:ss");
            }
            catch { return ""; }
        }

        public static void Release(object obj)
        {
            if (obj == null) return;
            try { Marshal.FinalReleaseComObject(obj); } catch { }
        }

        /// <summary>釋放後強制回收，避免 Outlook 程序殘留（Outlook COM 已知慣例）。</summary>
        public static void CoCleanup()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch { }
        }
    }
}
