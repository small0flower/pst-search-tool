using System;
using System.Threading;
using System.Windows.Forms;
using PstSearchTool.Diagnostics;

namespace PstSearchTool
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            StartupLog.Init();
            StartupLog.Log("程式進入 Main");

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                StartupLog.LogException("[AppDomain.UnhandledException]", e.ExceptionObject as Exception);
            Application.ThreadException += (s, e) =>
                StartupLog.LogException("[Application.ThreadException]", e.Exception);

            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                StartupLog.Log("建立主視窗…");
                var form = new UI.MainForm();
                StartupLog.Log("主視窗建立完成，進入訊息迴圈");
                Application.Run(form);
                StartupLog.Log("訊息迴圈結束");
            }
            catch (Exception ex)
            {
                StartupLog.LogException("[Main 未處理例外]", ex);
                throw;
            }
        }
    }
}
