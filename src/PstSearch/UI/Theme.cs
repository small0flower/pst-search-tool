using System.Drawing;
using System.Windows.Forms;

namespace PstSearchTool.UI
{
    /// <summary>深色/淺色主題：遞迴套用到整個控制項樹。</summary>
    internal static class Theme
    {
        public static bool IsDark;

        private static readonly Color DarkPanel = Color.FromArgb(37, 37, 38);
        private static readonly Color DarkControl = Color.FromArgb(45, 45, 48);
        private static readonly Color DarkFore = Color.FromArgb(225, 225, 225);
        private static readonly Color DarkBorder = Color.FromArgb(85, 85, 88);
        private static readonly Color DarkSel = Color.FromArgb(38, 79, 120);
        private static readonly Color DarkGrid = Color.FromArgb(60, 60, 62);
        private static readonly Color DarkHeader = Color.FromArgb(55, 55, 58);

        public static void Apply(Control root, bool dark)
        {
            IsDark = dark;
            ApplyRecursive(root, dark);
        }

        private static void ApplyRecursive(Control c, bool dark)
        {
            if (c is ToolStrip)
            {
                // 選單列保留系統樣式（避免自訂渲染器造成顯示異常）
            }
            else if (dark)
            {
                c.BackColor = DarkPanel;
                c.ForeColor = DarkFore;
                if (c is TextBoxBase || c is ComboBox || c is DateTimePicker || c is CheckedListBox || c is ListBox)
                {
                    c.BackColor = DarkControl;
                    c.ForeColor = DarkFore;
                }
                else if (c is TreeView tv)
                {
                    tv.BackColor = DarkControl;
                    tv.ForeColor = DarkFore;
                    tv.LineColor = DarkBorder;
                }
                else if (c is ListView lv)
                {
                    lv.BackColor = DarkControl;
                    lv.ForeColor = DarkFore;
                }
                else if (c is DataGridView g)
                {
                    g.BackgroundColor = DarkControl;
                    g.GridColor = DarkGrid;
                    g.DefaultCellStyle.BackColor = DarkControl;
                    g.DefaultCellStyle.ForeColor = DarkFore;
                    g.DefaultCellStyle.SelectionBackColor = DarkSel;
                    g.DefaultCellStyle.SelectionForeColor = Color.White;
                    g.ColumnHeadersDefaultCellStyle.BackColor = DarkHeader;
                    g.ColumnHeadersDefaultCellStyle.ForeColor = DarkFore;
                    g.EnableHeadersVisualStyles = false;
                    g.BorderStyle = BorderStyle.FixedSingle;
                }
                else if (c is Button b)
                {
                    b.FlatStyle = FlatStyle.Flat;
                    b.FlatAppearance.BorderColor = DarkBorder;
                    b.BackColor = DarkControl;
                    b.ForeColor = DarkFore;
                }
                else if (c is SplitContainer sc)
                {
                    sc.BackColor = DarkBorder;
                }
            }
            else
            {
                // 淺色：還原系統色
                c.BackColor = Color.Empty;
                c.ForeColor = Color.Empty;
                if (c is DataGridView g)
                {
                    g.EnableHeadersVisualStyles = true;
                    g.BackgroundColor = SystemColors.Window;
                    g.BorderStyle = BorderStyle.FixedSingle;
                }
                if (c is Button b)
                {
                    b.FlatStyle = FlatStyle.Standard;
                }
            }
            foreach (Control child in c.Controls) ApplyRecursive(child, dark);
        }
    }
}
