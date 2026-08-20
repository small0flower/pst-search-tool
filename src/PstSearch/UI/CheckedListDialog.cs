using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PstSearchTool.UI
{
    /// <summary>簡單的多選清單對話框（用於挑選要掛載的 .pst 檔案）。</summary>
    internal class CheckedListDialog : Form
    {
        private readonly CheckedListBox _list;

        public List<string> SelectedItems { get; private set; } = new List<string>();

        public CheckedListDialog(IEnumerable<string> items, string title, string prompt)
        {
            Text = title;
            Width = 560;
            Height = 460;
            MinimumSize = new System.Drawing.Size(420, 300);
            StartPosition = FormStartPosition.CenterParent;
            Font = new System.Drawing.Font("Microsoft JhengHei", 9F);

            var lbl = new Label
            {
                Text = prompt,
                Left = 12, Top = 10, Width = 520,
                AutoSize = false
            };

            _list = new CheckedListBox
            {
                Left = 12, Top = 40, Width = 520, Height = 320,
                CheckOnClick = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            foreach (var it in items) _list.Items.Add(it, true);

            var ok = new Button
            {
                Text = "確定", Left = 350, Top = 372, Width = 90,
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            var cancel = new Button
            {
                Text = "取消", Left = 450, Top = 372, Width = 90,
                DialogResult = DialogResult.Cancel,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };

            Controls.Add(lbl);
            Controls.Add(_list);
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
                foreach (var it in _list.CheckedItems) SelectedItems.Add(it.ToString());
            base.OnFormClosing(e);
        }
    }
}
