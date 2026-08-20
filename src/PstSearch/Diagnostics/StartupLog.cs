using System;
using System.IO;

namespace PstSearchTool.Diagnostics
{
    /// <summary>
    /// 啟動診斷日誌：把程式啟動過程的每一步寫入 %APPDATA%\PstSearchTool\startup.log。
    /// 用於定位「程式啟動但視窗不出現」類型的問題。
    /// </summary>
    internal static class StartupLog
    {
        private static readonly object Sync = new object();
        private static string _path;

        public static void Init()
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PstSearchTool");
                Directory.CreateDirectory(dir);
                _path = Path.Combine(dir, "startup.log");
                File.AppendAllText(_path,
                    Environment.NewLine + "===== 程式啟動 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " =====" + Environment.NewLine);
            }
            catch { }
        }

        public static void Log(string message)
        {
            try
            {
                if (_path == null) return;
                lock (Sync)
                {
                    File.AppendAllText(_path, DateTime.Now.ToString("HH:mm:ss.fff") + "  " + message + Environment.NewLine);
                }
            }
            catch { }
        }

        public static void LogException(string where, Exception ex)
        {
            Log(where + " 例外：" + (ex == null ? "(null)" : ex.GetType().FullName + ": " + ex.Message));
            if (ex != null && !string.IsNullOrEmpty(ex.StackTrace))
                Log(where + " StackTrace：" + ex.StackTrace);
        }
    }
}
