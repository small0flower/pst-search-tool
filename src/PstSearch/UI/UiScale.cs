using System;
using System.Drawing;
using System.Windows.Forms;

namespace PstSearchTool.UI
{
    /// <summary>
    /// DPI 縮放輔助：以 96 DPI 為設計基準，將程式化建置的 UI 依系統 DPI 等比放大，
    /// 避免高 DPI 顯示器（如 125%/150% 縮放）上文字過小。
    /// 字型透過表單 Font 的 ambient 繼承自動縮放，此處只處理位置/大小與欄寬。
    /// </summary>
    internal static class UiScale
    {
        public static float Factor(Control c)
        {
            try
            {
                using (var g = c.CreateGraphics())
                {
                    float s = g.DpiX / 96f;
                    return s < 1f ? 1f : s;
                }
            }
            catch { return 1f; }
        }

        public static Size ScaleSize(Size s, float scale)
        {
            return new Size((int)Math.Round(s.Width * scale), (int)Math.Round(s.Height * scale));
        }

        /// <summary>遞迴縮放控制項的位置與大小（字型已由表單 Font 繼承，不需重設）。</summary>
        public static void ScaleTree(Control parent, float scale)
        {
            if (scale <= 1.001f) return;
            foreach (Control c in parent.Controls)
            {
                c.Left = (int)Math.Round(c.Left * scale);
                c.Top = (int)Math.Round(c.Top * scale);
                c.Width = (int)Math.Round(c.Width * scale);
                c.Height = (int)Math.Round(c.Height * scale);
                if (c is ListView lv)
                {
                    foreach (ColumnHeader col in lv.Columns)
                        col.Width = (int)Math.Round(col.Width * scale);
                }
                else if (c is DataGridView g)
                {
                    foreach (DataGridViewColumn col in g.Columns)
                        col.Width = (int)Math.Round(col.Width * scale);
                    g.RowTemplate.Height = (int)Math.Round(g.RowTemplate.Height * scale);
                    g.ColumnHeadersHeight = (int)Math.Round(g.ColumnHeadersHeight * scale);
                }
                ScaleTree(c, scale);
            }
        }
    }
}
