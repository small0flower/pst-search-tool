using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using PstSearchTool.Config;
using PstSearchTool.Data;
using PstSearchTool.Diagnostics;
using PstSearchTool.Indexing;
using PstSearchTool.Models;
using PstSearchTool.Outlook;
using PstSearchTool.Search;

namespace PstSearchTool.UI
{
    /// <summary>主視窗：來源選擇 → 資料夾勾選 → 建立索引 → 搜尋 → 於 Outlook 開啟。</summary>
    internal class MainForm : Form
    {
        // 頂端來源列
        private RadioButton _rbCurrentPst;
        private RadioButton _rbExternalPst;
        private TextBox _txtDir;
        private Button _btnBrowseDir;
        private Button _btnRefreshStores;
        // 左側
        private ListView _lvStores;
        private TreeView _tvFolders;
        private Button _btnInbox;
        private Button _btnSent;
        private Button _btnAllFolders;
        private Button _btnClearFolders;
        // 右側搜尋條件
        private TextBox _txtKeyword;
        private Button _btnSearch;
        private Button _btnClearSearch;
        private CheckBox _chkDate;
        private DateTimePicker _dtFrom;
        private DateTimePicker _dtTo;
        private TextBox _txtSender;
        private ComboBox _cmbFolderFilter;
        private DataGridView _grid;
        // 底部
        private Button _btnIndex;
        private Button _btnRebuild;
        private Button _btnCancel;
        private ProgressBar _progress;
        private Label _lblStatus;

        private readonly AppSettings _settings = Settings.Load();
        private IndexStore _db;
        private CancellationTokenSource _cts;
        private List<StoreInfo> _stores = new List<StoreInfo>();
        private bool _loadingTree;
        private bool _busy;
        private bool _suppressSource;
        private string _currentTreeStoreId = "";
        private SplitContainer _sc;
        private float _dpiScale = 1f;

        public MainForm()
        {
            StartupLog.Log("MainForm 建構開始");
            _dpiScale = UiScale.Factor(this);
            Font = new Font("Microsoft JhengHei", 9.5f * _dpiScale);
            Size = UiScale.ScaleSize(new Size(1180, 760), _dpiScale);
            MinimumSize = UiScale.ScaleSize(new Size(940, 620), _dpiScale);
            BuildUi();
            UiScale.ScaleTree(this, _dpiScale);
            StartupLog.Log("BuildUi 完成（DPI 縮放：" + _dpiScale.ToString("0.00") + "）");
            Load += OnLoad;
            Shown += OnShown;
            FormClosing += OnFormClosing;
        }

        // ------------------------------------------------------------------ UI 建置
        private void BuildUi()
        {
            Text = "Outlook PST 搜尋工具";
            Font = new Font("Microsoft JhengHei", 9F);
            Size = new Size(1180, 760);
            MinimumSize = new Size(940, 620);
            StartPosition = FormStartPosition.CenterScreen;
            KeyPreview = true;

            // --- 頂端來源列 ---
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 62 };
            _rbCurrentPst = new RadioButton { Text = "目前 Outlook 中的 PST", Left = 10, Top = 10, AutoSize = true };
            _rbExternalPst = new RadioButton { Text = "外掛 PST", Left = 260, Top = 10, AutoSize = true };
            _txtDir = new TextBox { Left = 370, Top = 8, Width = 400, ReadOnly = true };
            _btnBrowseDir = new Button { Text = "瀏覽…", Left = 778, Top = 6, Width = 76 };
            _btnRefreshStores = new Button { Text = "重新整理來源", Left = 862, Top = 6, Width = 110 };
            var lblHint = new Label
            {
                Text = "外掛 PST 會加入 Outlook 設定檔以建立索引與開啟郵件（之後可於 Outlook「檔案→資料檔管理」移除）。",
                Left = 10, Top = 36, AutoSize = true, ForeColor = Color.Gray
            };
            pnlTop.Controls.AddRange(new Control[] { _rbCurrentPst, _rbExternalPst, _txtDir, _btnBrowseDir, _btnRefreshStores, lblHint });

            // --- 主分割（分隔距離與最小尺寸需於表單配置完成後設定，見 OnLoad）---
            _sc = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical };
            _sc.SplitterWidth = 5;

            // ==== 左側：存放區 + 資料夾 ====
            var lblStores = new Label { Text = "2. 郵件來源（勾選要索引的）", Left = 10, Top = 8, AutoSize = true, Font = new Font(Font, FontStyle.Bold) };
            _lvStores = new ListView
            {
                Left = 10, Top = 32, Width = 360, Height = 150,
                View = View.Details, CheckBoxes = true, FullRowSelect = true, MultiSelect = false, HideSelection = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _lvStores.Columns.Add("名稱", 150);
            _lvStores.Columns.Add("檔案", 200);

            var lblFolders = new Label { Text = "3. 資料夾（勾選要索引的；含其子資料夾）", Left = 10, Top = 192, AutoSize = true, Font = new Font(Font, FontStyle.Bold) };
            _btnInbox = new Button { Text = "收件匣", Left = 10, Top = 214, Width = 72 };
            _btnSent = new Button { Text = "寄件匣", Left = 90, Top = 214, Width = 72 };
            _btnAllFolders = new Button { Text = "全選", Left = 170, Top = 214, Width = 60 };
            _btnClearFolders = new Button { Text = "清空", Left = 238, Top = 214, Width = 60 };
            _tvFolders = new TreeView
            {
                Left = 10, Top = 246, Width = 360, CheckBoxes = true, HideSelection = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            _sc.Panel1.Controls.AddRange(new Control[] { lblStores, _lvStores, lblFolders, _btnInbox, _btnSent, _btnAllFolders, _btnClearFolders, _tvFolders });
            // ==== 右側：搜尋 + 結果 ====
            var pnlSearch = new Panel { Dock = DockStyle.Top, Height = 118 };
            var lblKw = new Label { Text = "4. 搜尋條件", Left = 10, Top = 8, AutoSize = true, Font = new Font(Font, FontStyle.Bold) };
            _txtKeyword = new TextBox { Left = 10, Top = 32, Width = 380 };
            _btnSearch = new Button { Text = "搜尋", Left = 400, Top = 30, Width = 90 };
            _btnClearSearch = new Button { Text = "清除", Left = 498, Top = 30, Width = 70 };
            _chkDate = new CheckBox { Text = "日期：", Left = 10, Top = 64, AutoSize = true };
            _dtFrom = new DateTimePicker { Left = 60, Top = 62, Width = 120, Format = DateTimePickerFormat.Short };
            var lblDash = new Label { Text = "～", Left = 184, Top = 66, AutoSize = true };
            _dtTo = new DateTimePicker { Left = 204, Top = 62, Width = 120, Format = DateTimePickerFormat.Short };
            var lblSender = new Label { Text = "寄件者：", Left = 10, Top = 92, AutoSize = true };
            _txtSender = new TextBox { Left = 70, Top = 90, Width = 220 };
            var lblFolderF = new Label { Text = "資料夾：", Left = 310, Top = 92, AutoSize = true };
            _cmbFolderFilter = new ComboBox { Left = 370, Top = 90, Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbFolderFilter.Items.AddRange(new object[] { "全部資料夾", "收件匣", "寄件匣", "自訂（左側勾選）" });
            pnlSearch.Controls.AddRange(new Control[] { lblKw, _txtKeyword, _btnSearch, _btnClearSearch, _chkDate, _dtFrom, lblDash, _dtTo, lblSender, _txtSender, lblFolderF, _cmbFolderFilter });

            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            _btnIndex = new Button { Text = "建立/更新索引", Left = 10, Top = 9, Width = 120 };
            _btnRebuild = new Button { Text = "完整重建索引", Left = 140, Top = 9, Width = 110 };
            _btnCancel = new Button { Text = "取消", Left = 260, Top = 9, Width = 70, Enabled = false };
            _progress = new ProgressBar { Left = 345, Top = 11, Width = 280 };
            _lblStatus = new Label { Left = 640, Top = 13, Width = 520, AutoEllipsis = true, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            pnlBottom.Controls.AddRange(new Control[] { _btnIndex, _btnRebuild, _btnCancel, _progress, _lblStatus });

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                BackgroundColor = SystemColors.Window,
                AllowUserToResizeRows = false
            };
            var c0 = new DataGridViewTextBoxColumn { HeaderText = "日期", Width = 110, SortMode = DataGridViewColumnSortMode.Automatic };
            var c1 = new DataGridViewTextBoxColumn { HeaderText = "寄件者", Width = 160, SortMode = DataGridViewColumnSortMode.Automatic };
            var c2 = new DataGridViewTextBoxColumn { HeaderText = "主旨", Width = 260, SortMode = DataGridViewColumnSortMode.Automatic };
            var c3 = new DataGridViewTextBoxColumn { HeaderText = "資料夾", Width = 130, SortMode = DataGridViewColumnSortMode.Automatic };
            var c4 = new DataGridViewTextBoxColumn { HeaderText = "摘要", Width = 320, SortMode = DataGridViewColumnSortMode.NotSortable };
            var c5 = new DataGridViewTextBoxColumn { HeaderText = "存放區", Width = 90, SortMode = DataGridViewColumnSortMode.NotSortable };
            _grid.Columns.AddRange(new DataGridViewColumn[] { c0, c1, c2, c3, c4, c5 });

            _sc.Panel2.Controls.Add(_grid);
            _sc.Panel2.Controls.Add(pnlSearch);
            _sc.Panel2.Controls.Add(pnlBottom);

            Controls.Add(_sc);
            Controls.Add(pnlTop);

            // --- 事件 ---
            _rbCurrentPst.CheckedChanged += SourceModeChanged;
            _rbExternalPst.CheckedChanged += SourceModeChanged;
            _btnBrowseDir.Click += BtnBrowseDir_Click;
            _btnRefreshStores.Click += (s, e) => RefreshStores();
            _lvStores.ItemSelectionChanged += LvStores_ItemSelectionChanged;
            _tvFolders.AfterCheck += TvFolders_AfterCheck;
            _btnInbox.Click += (s, e) => CheckKindOnly("inbox");
            _btnSent.Click += (s, e) => CheckKindOnly("sent");
            _btnAllFolders.Click += (s, e) => SetAllChecked(true);
            _btnClearFolders.Click += (s, e) => SetAllChecked(false);
            _btnIndex.Click += (s, e) => StartIndex(false);
            _btnRebuild.Click += (s, e) => StartIndex(true);
            _btnCancel.Click += (s, e) => { try { _cts?.Cancel(); } catch { } _lblStatus.Text = "正在取消…"; };
            _btnSearch.Click += (s, e) => DoSearch();
            _btnClearSearch.Click += (s, e) =>
            {
                _txtKeyword.Clear();
                _txtSender.Clear();
                _chkDate.Checked = false;
                _cmbFolderFilter.SelectedIndex = 0;
            };
            _txtKeyword.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; DoSearch(); } };
            _grid.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) OpenSelected(); };
            _grid.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter && _grid.SelectedRows.Count > 0) { e.Handled = true; OpenSelected(); } };
        }

        // ------------------------------------------------------------------ 生命週期
        private void OnLoad(object sender, EventArgs e)
        {
            StartupLog.Log("OnLoad 開始");
            try
            {
                StartupLog.Log("開啟索引資料庫：" + _settings.DbPath);
                _db = new IndexStore(_settings.DbPath);
                StartupLog.Log("索引資料庫開啟成功");
            }
            catch (Exception ex)
            {
                StartupLog.LogException("開啟索引資料庫失敗", ex);
                MessageBox.Show(this, "無法開啟索引資料庫：\n" + ex.Message + "\n\n請確認磁碟空間與寫入權限。", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                _btnIndex.Enabled = _btnRebuild.Enabled = _btnSearch.Enabled = false;
            }

            try
            {
                if (_settings.WindowW > 0)
                {
                    StartPosition = FormStartPosition.Manual;
                    SetBounds(_settings.WindowX, _settings.WindowY, _settings.WindowW, _settings.WindowH, BoundsSpecified.All);
                    if (_settings.WindowMaximized) WindowState = FormWindowState.Maximized;
                }
                // 視窗不得超出螢幕工作區（高 DPI 縮放後可能超過）
                try
                {
                    var wa = Screen.FromControl(this).WorkingArea;
                    if (Width > wa.Width - 20) Width = wa.Width - 20;
                    if (Height > wa.Height - 40) Height = wa.Height - 40;
                    if (Left < wa.Left) Left = wa.Left;
                    if (Top < wa.Top) Top = wa.Top;
                }
                catch { }

                // SplitContainer 需於表單完成配置後才設定分隔距離（建構時寬度為 0 會拋例外）
                try
                {
                    _sc.SplitterDistance = (int)Math.Round(380 * _dpiScale);
                    _sc.Panel1MinSize = (int)Math.Round(300 * _dpiScale);
                    _sc.Panel2MinSize = (int)Math.Round(520 * _dpiScale);
                }
                catch (Exception ex)
                {
                    StartupLog.LogException("設定 SplitContainer 失敗", ex);
                }

                _suppressSource = true;
                if (_settings.SourceMode == 1) _rbExternalPst.Checked = true;
                else _rbCurrentPst.Checked = true;
                _suppressSource = false;
                UpdateSourceControls();

                _cmbFolderFilter.SelectedIndex = 0;
                _txtKeyword.Text = _settings.LastKeyword;
                _txtSender.Text = _settings.SenderFilter;
                _chkDate.Checked = _settings.DateFilterEnabled;
                DateTime df, dt;
                if (DateTime.TryParse(_settings.DateFrom, out df) && df != DateTime.MinValue) _dtFrom.Value = df;
                else _dtFrom.Value = DateTime.Today.AddMonths(-3);
                if (DateTime.TryParse(_settings.DateTo, out dt) && dt != DateTime.MinValue) _dtTo.Value = dt;
                else _dtTo.Value = DateTime.Today;

                if (_db != null) _lblStatus.Text = "已索引 " + _db.CountMessages() + " 封郵件。";
                StartupLog.Log("OnLoad 完成");
            }
            catch (Exception ex)
            {
                StartupLog.LogException("OnLoad 後段例外", ex);
            }
        }

        /// <summary>視窗已顯示後才載入郵件來源（避免 COM 呼叫阻塞視窗顯示）。</summary>
        private void OnShown(object sender, EventArgs e)
        {
            StartupLog.Log("視窗已顯示，開始載入郵件來源");
            RefreshStores();
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            StartupLog.Log("程式關閉中");
            SaveCurrentTreeChecks();
            _settings.LastKeyword = _txtKeyword.Text;
            _settings.SenderFilter = _txtSender.Text;
            _settings.DateFilterEnabled = _chkDate.Checked;
            _settings.DateFrom = _dtFrom.Value.ToString("yyyy-MM-dd");
            _settings.DateTo = _dtTo.Value.ToString("yyyy-MM-dd");
            _settings.SourceMode = _rbExternalPst.Checked ? 1 : 0;
            _settings.ExternalDir = _txtDir.Text;
            if (WindowState == FormWindowState.Normal)
            {
                _settings.WindowX = Location.X;
                _settings.WindowY = Location.Y;
                _settings.WindowW = Size.Width;
                _settings.WindowH = Size.Height;
                _settings.WindowMaximized = false;
            }
            else _settings.WindowMaximized = (WindowState == FormWindowState.Maximized);
            Settings.Save(_settings);
            try { _cts?.Cancel(); } catch { }
            if (_db != null) { try { _db.Dispose(); } catch { } }
        }

        // ------------------------------------------------------------------ 來源與資料夾
        private void SourceModeChanged(object sender, EventArgs e)
        {
            if (_suppressSource) return;
            UpdateSourceControls();
            RefreshStores();
        }

        private void UpdateSourceControls()
        {
            bool ext = _rbExternalPst.Checked;
            _txtDir.Enabled = ext;
            _btnBrowseDir.Enabled = ext;
        }

        private void BtnBrowseDir_Click(object sender, EventArgs e)
        {
            if (!_rbExternalPst.Checked) { _rbExternalPst.Checked = true; return; }
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "選擇包含 .pst 檔案的資料夾（會掃描子資料夾）";
                dlg.SelectedPath = _settings.ExternalDir;
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                _txtDir.Text = dlg.SelectedPath;
                _settings.ExternalDir = dlg.SelectedPath;
            }
            List<string> psts;
            try
            {
                psts = Directory.GetFiles(_txtDir.Text, "*.pst", SearchOption.AllDirectories).OrderBy(x => x).ToList();
            }
            catch (Exception ex) { Err("讀取資料夾失敗：" + ex.Message); return; }
            if (psts.Count == 0) { Err("此資料夾中找不到 .pst 檔案。"); return; }

            using (var dlg = new CheckedListDialog(psts, "選擇要掛載的 PST 檔案", "勾選要加入 Outlook 並建立索引的 .pst 檔案："))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                if (dlg.SelectedItems.Count == 0) return;
                if (MessageBox.Show(this,
                    "將把下列 " + dlg.SelectedItems.Count + " 個 PST 加入 Outlook 設定檔以建立索引：\n\n" +
                    string.Join("\n", dlg.SelectedItems.Select(Path.GetFileName)) +
                    "\n\n加入後它們會出現在 Outlook 的資料夾清單中，可隨時於 Outlook「檔案→資料檔管理」移除。是否繼續？",
                    "確認掛載", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;

                var files = dlg.SelectedItems.ToList();
                SetBusy(true, "正在掛載 PST…");
                Task.Run(() =>
                {
                    var errs = new List<string>();
                    foreach (var f in files)
                    {
                        try { OutlookService.AddStore(f); }
                        catch (Exception ex) { errs.Add(Path.GetFileName(f) + "：" + ex.Message); }
                    }
                    Ui(() =>
                    {
                        _settings.ExternalPstFiles = files;
                        if (errs.Count > 0) Err("部分 PST 掛載失敗：\n" + string.Join("\n", errs));
                        RefreshStores();
                    });
                });
            }
        }

        private void RefreshStores()
        {
            StartupLog.Log("RefreshStores 開始（模式：" + (_rbExternalPst.Checked ? "外掛 PST" : "目前 PST") + "）");
            SetBusy(true, "正在讀取 Outlook 郵件來源…");
            Task.Run(() =>
            {
                try
                {
                    List<StoreInfo> all = OutlookService.GetStores();
                    StartupLog.Log("GetStores 回傳 " + all.Count + " 個存放區");
                    List<StoreInfo> shown;
                    if (_rbExternalPst.Checked && _settings.ExternalPstFiles.Count > 0)
                        shown = all.Where(s => _settings.ExternalPstFiles.Any(f => PathEquals(f, s.FilePath))).ToList();
                    else if (_rbExternalPst.Checked)
                        shown = new List<StoreInfo>();
                    else
                        shown = all;

                    Ui(() =>
                    {
                        _stores = shown;
                        _lvStores.BeginUpdate();
                        _lvStores.Items.Clear();
                        foreach (var s in shown)
                        {
                            var li = new ListViewItem(s.DisplayName) { Tag = s, Checked = IsStoreChecked(s.StoreId) };
                            li.SubItems.Add(s.FilePath);
                            _lvStores.Items.Add(li);
                        }
                        _lvStores.EndUpdate();
                        if (shown.Count == 0 && _rbExternalPst.Checked)
                            _lblStatus.Text = "尚未選擇外掛 PST（請按「瀏覽…」選取 .pst 檔案）。";
                        else if (shown.Count == 0)
                            _lblStatus.Text = "Outlook 設定檔中沒有可用的郵件存放區。";
                        SetBusy(false, "就緒");
                        SelectFirstCheckedStore();
                    });
                }
                catch (Exception ex)
                {
                    StartupLog.LogException("RefreshStores 失敗", ex);
                    Ui(() => { Err("讀取 Outlook 失敗：" + ex.Message); SetBusy(false, "就緒"); });
                }
            });
        }

        private bool IsStoreChecked(string storeId)
        {
            if (_settings.Stores.Count == 0)
            {
                var s = _stores.FirstOrDefault(x => x.StoreId == storeId);
                return s != null && s.IsDefault;
            }
            return _settings.Stores.Any(x => x.StoreId == storeId);
        }

        private void SelectFirstCheckedStore()
        {
            foreach (ListViewItem li in _lvStores.Items)
                if (li.Checked) { li.Selected = true; break; }
        }

        private void LvStores_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (!e.Item.Selected || _busy) return;
            var store = e.Item.Tag as StoreInfo;
            if (store == null || store.StoreId == _currentTreeStoreId) return;
            SaveCurrentTreeChecks();
            LoadTree(store);
        }

        private void LoadTree(StoreInfo store)
        {
            _tvFolders.Nodes.Clear();
            _currentTreeStoreId = store.StoreId;
            _lblStatus.Text = "正在讀取資料夾…";
            string storeId = store.StoreId;
            Task.Run(() =>
            {
                List<FolderNode> tree = null;
                string err = null;
                try { tree = OutlookService.GetFolderTree(storeId); }
                catch (Exception ex) { err = ex.Message; }
                Ui(() =>
                {
                    if (_currentTreeStoreId != storeId) return;
                    _loadingTree = true;
                    _tvFolders.BeginUpdate();
                    _tvFolders.Nodes.Clear();
                    if (tree != null)
                    {
                        var st = _settings.Stores.FirstOrDefault(x => x.StoreId == storeId);
                        HashSet<string> sel = st == null
                            ? null
                            : new HashSet<string>(st.SelectedFolders.Select(DecodeFolder).Select(f => f.Path));
                        foreach (var fn in tree) _tvFolders.Nodes.Add(BuildNode(fn, sel));
                        if (sel == null) { CheckKind("inbox"); CheckKind("sent"); }
                        _tvFolders.ExpandAll();
                    }
                    _tvFolders.EndUpdate();
                    _loadingTree = false;
                    _lblStatus.Text = err != null ? "讀取資料夾失敗：" + err : "資料夾已載入。";
                });
            });
        }

        private TreeNode BuildNode(FolderNode fn, HashSet<string> sel)
        {
            var tn = new TreeNode(fn.Name) { Tag = fn };
            if (sel != null && sel.Contains(fn.Path)) tn.Checked = true;
            foreach (var c in fn.Children) tn.Nodes.Add(BuildNode(c, sel));
            return tn;
        }

        private void TvFolders_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (!_loadingTree) SaveCurrentTreeChecks();
        }

        private void CheckKindOnly(string kind)
        {
            if (_tvFolders.Nodes.Count == 0) return;
            _loadingTree = true;
            SetAllChecked(false);
            CheckKind(kind);
            _loadingTree = false;
            SaveCurrentTreeChecks();
        }

        private void SetAllChecked(bool ck)
        {
            if (_tvFolders.Nodes.Count == 0) return;
            _loadingTree = true;
            SetAllCheckedRec(_tvFolders.Nodes, ck);
            _loadingTree = false;
            SaveCurrentTreeChecks();
        }

        private void SetAllCheckedRec(TreeNodeCollection nodes, bool ck)
        {
            foreach (TreeNode n in nodes)
            {
                n.Checked = ck;
                SetAllCheckedRec(n.Nodes, ck);
            }
        }

        private void CheckKind(string kind)
        {
            foreach (TreeNode n in _tvFolders.Nodes) CheckKindRec(n, kind);
        }

        private void CheckKindRec(TreeNode n, string kind)
        {
            var fn = n.Tag as FolderNode;
            if (fn != null && fn.Kind == kind) n.Checked = true;
            foreach (TreeNode c in n.Nodes) CheckKindRec(c, kind);
        }

        private List<FolderInfo> GetCheckedFolderPaths()
        {
            var result = new List<FolderInfo>();
            CollectChecked(_tvFolders.Nodes, result);
            return result;
        }

        private void CollectChecked(TreeNodeCollection nodes, List<FolderInfo> list)
        {
            foreach (TreeNode n in nodes)
            {
                var fn = n.Tag as FolderNode;
                if (fn != null && n.Checked)
                {
                    list.Add(new FolderInfo { Path = fn.Path, Kind = fn.Kind });
                    AddDescendants(fn, list);
                }
                else CollectChecked(n.Nodes, list);
            }
        }

        private void AddDescendants(FolderNode fn, List<FolderInfo> list)
        {
            foreach (var c in fn.Children)
            {
                list.Add(new FolderInfo { Path = c.Path, Kind = c.Kind });
                AddDescendants(c, list);
            }
        }

        private void SaveCurrentTreeChecks()
        {
            if (_tvFolders.Nodes.Count == 0) return;
            var store = CurrentStore();
            if (store == null) return;
            var sel = GetCheckedFolderPaths().Select(EncodeFolder).ToList();
            var st = _settings.Stores.FirstOrDefault(x => x.StoreId == store.StoreId);
            if (st == null)
            {
                st = new StoreSetting { StoreId = store.StoreId, DisplayName = store.DisplayName };
                _settings.Stores.Add(st);
            }
            st.SelectedFolders = sel;
        }

        private StoreInfo CurrentStore()
        {
            if (_lvStores.SelectedItems.Count > 0) return _lvStores.SelectedItems[0].Tag as StoreInfo;
            return null;
        }

        private static string EncodeFolder(FolderInfo f) => f.Kind + "::" + f.Path;

        private static FolderInfo DecodeFolder(string s)
        {
            int i = s.IndexOf("::", StringComparison.Ordinal);
            if (i < 0) return new FolderInfo { Kind = "other", Path = s };
            return new FolderInfo { Kind = s.Substring(0, i), Path = s.Substring(i + 2) };
        }

        private List<FolderInfo> GetSavedFolderSelection(string storeId)
        {
            var st = _settings.Stores.FirstOrDefault(x => x.StoreId == storeId);
            if (st == null) return new List<FolderInfo>();
            return st.SelectedFolders.Select(DecodeFolder).ToList();
        }

        // ------------------------------------------------------------------ 索引
        private void StartIndex(bool rebuild)
        {
            if (_db == null) { Err("索引資料庫不可用。"); return; }
            SaveCurrentTreeChecks();
            var stores = GetCheckedStores();
            if (stores.Count == 0) { Err("請先在「郵件來源」勾選至少一個存放區。"); return; }

            var plan = new List<KeyValuePair<StoreInfo, List<FolderInfo>>>();
            var skipped = new List<string>();
            foreach (var s in stores)
            {
                var sel = GetSavedFolderSelection(s.StoreId);
                if (sel.Count == 0) skipped.Add(s.DisplayName);
                else plan.Add(new KeyValuePair<StoreInfo, List<FolderInfo>>(s, sel));
            }
            if (plan.Count == 0)
            {
                Err("所有勾選的存放區都尚未選取資料夾。\n請依序點選每個存放區，並在左側資料夾樹勾選（如收件匣/寄件匣）。");
                return;
            }

            if (!_settings.HintShown)
            {
                MessageBox.Show(this,
                    "首次建立索引前的小提醒：\n\n" +
                    "1. 索引期間請不要關閉 Outlook；\n" +
                    "2. 若 Outlook 彈出「程式嘗試存取…」安全性警告，請點選「允許」；\n" +
                    "3. 大型 PST 首次索引可能需要一段時間，可隨時按「取消」。",
                    "開始索引", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _settings.HintShown = true;
            }

            string skipMsg = skipped.Count > 0
                ? "\n（已略過未選資料夾的存放區：" + string.Join("、", skipped) + "）"
                : "";
            if (MessageBox.Show(this,
                "開始建立/更新索引？（" + plan.Count + " 個存放區）" + skipMsg,
                "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;

            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            SetBusy(true, "正在建立索引…", true);
            var db = _db;
            Task.Run(() =>
            {
                long total = 0;
                try
                {
                    var indexer = new Indexer(db);
                    indexer.StatusChanged += (fp, msg) => Ui(() => _lblStatus.Text = "[" + fp + "] " + msg);
                    indexer.FolderProgress += (fp, done, all) => Ui(() =>
                    {
                        if (all > 0)
                        {
                            _progress.Style = ProgressBarStyle.Blocks;
                            _progress.Maximum = (int)Math.Min(all, int.MaxValue);
                            _progress.Value = (int)Math.Min(done, int.MaxValue);
                        }
                        else _progress.Style = ProgressBarStyle.Marquee;
                    });
                    foreach (var kv in plan)
                    {
                        if (token.IsCancellationRequested) break;
                        total += indexer.Run(kv.Key, kv.Value, rebuild, () => token.IsCancellationRequested);
                    }
                }
                catch (Exception ex)
                {
                    Ui(() => Err("索引失敗：" + ex.Message));
                }
                finally
                {
                    string final = token.IsCancellationRequested
                        ? "索引已取消（已完成的資料夾已儲存）。"
                        : (total > 0 ? "索引完成，共 " + total + " 封。" : "沒有可索引的郵件。");
                    Ui(() => { SetBusy(false, ""); _lblStatus.Text = final; });
                    try { _cts.Dispose(); } catch { }
                    _cts = null;
                }
            });
        }

        private List<StoreInfo> GetCheckedStores()
        {
            return _lvStores.CheckedItems.Cast<ListViewItem>().Select(li => li.Tag as StoreInfo).Where(s => s != null).ToList();
        }

        // ------------------------------------------------------------------ 搜尋
        private void DoSearch()
        {
            if (_db == null) { Err("索引資料庫不可用。"); return; }
            SaveCurrentTreeChecks();
            string kw = _txtKeyword.Text.Trim();
            var keywords = kw.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            string folderKind = null;
            List<string> folderPaths = null;
            switch (_cmbFolderFilter.SelectedIndex)
            {
                case 1: folderKind = "inbox"; break;
                case 2: folderKind = "sent"; break;
                case 3:
                    folderPaths = GetCheckedFolderPaths().Select(f => f.Path).Distinct().ToList();
                    if (folderPaths.Count == 0) { Err("請先在左側資料夾樹勾選要搜尋的資料夾（自訂模式）。"); return; }
                    break;
            }
            DateTime? from = _chkDate.Checked ? (DateTime?)_dtFrom.Value.Date : null;
            DateTime? to = _chkDate.Checked ? (DateTime?)_dtTo.Value.Date : null;
            string sender = _txtSender.Text.Trim();
            if (keywords.Length == 0 && sender.Length == 0 && !from.HasValue && folderKind == null && folderPaths == null)
            { Err("請輸入至少一個搜尋條件（關鍵字、寄件者或日期）。"); return; }

            SetBusy(true, "搜尋中…");
            var db = _db;
            Task.Run(() =>
            {
                try
                {
                    var results = db.Search(keywords, folderPaths, from, to, sender, folderKind, 500);
                    Ui(() =>
                    {
                        PopulateGrid(results);
                        SetBusy(false, "");
                    });
                }
                catch (Exception ex)
                {
                    Ui(() => { Err("搜尋失敗：" + ex.Message); SetBusy(false, "就緒"); });
                }
            });
        }

        private void PopulateGrid(List<SearchResultItem> results)
        {
            _grid.Rows.Clear();
            foreach (var r in results)
            {
                int idx = _grid.Rows.Add();
                var row = _grid.Rows[idx];
                row.Cells[0].Value = r.ReceivedTime;
                row.Cells[1].Value = r.FromName + (string.IsNullOrEmpty(r.FromEmail) ? "" : " <" + r.FromEmail + ">");
                row.Cells[2].Value = r.Subject;
                row.Cells[3].Value = r.FolderPath;
                row.Cells[4].Value = r.Snippet;
                row.Cells[5].Value = r.StoreName;
                row.Tag = r;
            }
            _lblStatus.Text = "找到 " + results.Count + " 筆" + (results.Count >= 500 ? "（已達顯示上限 500 筆）" : "") + "。雙擊即可於 Outlook 開啟。";
        }

        // ------------------------------------------------------------------ 開啟郵件
        private void OpenSelected()
        {
            if (_grid.SelectedRows.Count == 0) return;
            var item = _grid.SelectedRows[0].Tag as SearchResultItem;
            if (item == null) return;
            SetBusy(true, "正在於 Outlook 開啟…");
            Task.Run(() =>
            {
                try { OutlookService.OpenItem(item.EntryId, item.StoreId); }
                catch (Exception ex) { Ui(() => Err("無法開啟郵件：" + ex.Message)); }
                finally { Ui(() => SetBusy(false, "就緒")); }
            });
        }

        // ------------------------------------------------------------------ 工具
        private static bool PathEquals(string a, string b)
        {
            try
            {
                return string.Equals(Path.GetFullPath(a ?? ""), Path.GetFullPath(b ?? ""), StringComparison.OrdinalIgnoreCase);
            }
            catch { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }
        }

        private void SetBusy(bool busy, string status, bool indexing = false)
        {
            _busy = busy;
            bool ok = !busy && _db != null;
            _btnIndex.Enabled = ok;
            _btnRebuild.Enabled = ok;
            _btnSearch.Enabled = ok;
            _btnCancel.Enabled = indexing && busy;
            _btnBrowseDir.Enabled = !busy && _rbExternalPst.Checked;
            _btnRefreshStores.Enabled = !busy;
            _rbCurrentPst.Enabled = !busy;
            _rbExternalPst.Enabled = !busy;
            _lvStores.Enabled = !busy;
            _tvFolders.Enabled = !busy;
            _btnInbox.Enabled = !busy;
            _btnSent.Enabled = !busy;
            _btnAllFolders.Enabled = !busy;
            _btnClearFolders.Enabled = !busy;
            _txtKeyword.Enabled = !busy;
            _txtSender.Enabled = !busy;
            _chkDate.Enabled = !busy;
            _dtFrom.Enabled = !busy;
            _dtTo.Enabled = !busy;
            _cmbFolderFilter.Enabled = !busy;
            _btnClearSearch.Enabled = !busy;
            if (!string.IsNullOrEmpty(status)) _lblStatus.Text = status;
            if (!busy) { _progress.Style = ProgressBarStyle.Blocks; _progress.Value = 0; }
        }

        private void Ui(Action a)
        {
            if (IsDisposed) return;
            try
            {
                if (InvokeRequired) BeginInvoke(a);
                else a();
            }
            catch { }
        }

        private void Err(string text)
        {
            MessageBox.Show(this, text, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
