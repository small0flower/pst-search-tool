using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PstSearchTool.Models;

namespace PstSearchTool.UI
{
    /// <summary>搜尋統計：依寄件者 / 資料夾 / 月份。</summary>
    internal class StatsDialog : Form
    {
        public StatsDialog(StatsResult stats)
        {
            float s = UiScale.Factor(this);
            Text = "搜尋統計";
            Width = (int)Math.Round(790 * s);
            Height = (int)Math.Round(560 * s);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft JhengHei", 9F * s);

            var lblTotal = new Label
            {
                Text = "已索引郵件總數：" + stats.Total,
                Left = 12, Top = 8, AutoSize = true, Font = new Font(Font, FontStyle.Bold)
            };
            Controls.Add(lblTotal);

            AddGroup("依寄件者（Top 20）", 12, 44, stats.TopSenders);
            AddGroup("依資料夾", 272, 44, stats.FolderCounts);
            AddGroup("依月份", 532, 44, stats.MonthCounts);

            var ok = new Button
            {
                Text = "關閉",
                Left = (int)Math.Round(350 * s),
                Top = (int)Math.Round(505 * s),
                Width = 90,
                DialogResult = DialogResult.OK
            };
            Controls.Add(ok);
            AcceptButton = ok;
            CancelButton = ok;
            UiScale.ScaleTree(this, s);
        }

        private void AddGroup(string title, int left, int top, List<KeyValuePair<string, long>> data)
        {
            var lbl = new Label
            {
                Text = title, Left = left, Top = top, AutoSize = true, Font = new Font(Font, FontStyle.Bold)
            };
            var lv = new ListView
            {
                Left = left, Top = top + 22, Width = 245, Height = 425,
                View = View.Details, HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            lv.Columns.Add("項目", 155);
            lv.Columns.Add("數量", 70);
            foreach (var kv in data)
            {
                var it = new ListViewItem(kv.Key);
                it.SubItems.Add(kv.Value.ToString());
                lv.Items.Add(it);
            }
            Controls.Add(lbl);
            Controls.Add(lv);
        }
    }
}
