using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace PstSearchTool
{
    /// <summary>
    /// 多語系：目前僅支援「繁體 → 簡體」的介面字元轉換。
    /// 採用集中在載入時對「介面鍍鉻」做一次繁→簡轉換，不更動版面與邏輯；
    /// 資料控制項（資料夾樹節點、清單項目、表格儲存格＝郵件內容）一律不轉換。
    /// </summary>
    internal static class Lang
    {
        public static bool Simplified;   // 是否簡體

        // 繁體 → 簡體字元對照（僅列相異字元，依序對應）
        private const string TString = "並併來係個們側備傳僅儲內兩刪則動匯區協問啟單嘗嚴圖執塊夾實寫寬將專尋對導層屬庫強彈後徑從態慣應戶拋掃掛採換擇擊擬擴擷敗數"
            + "斷於時會條構標樹檔欄權殘沒況減測準滾滿濾瀏為無狀現環產畫異當疊發確稱筆節簡終組結給統緒線縮總繫繼續脫與舊蓋處號螢補裝製複"
            + "見視覽觸訂計訊記設許診詞詢試話該誌認語誤說調請證護讀變讓負責貼資賴較載輔輪輯輸迴這連週進運過達遞遷選還邊邏郵鄰釋鈕錯鍵長"
            + "閉開間關際隨隱雙離電頂項順須預題類顯餘驗體鴻點齊";

        private const string SString = "并并来系个们侧备传仅储内两删则动汇区协问启单尝严图执块夹实写宽将专寻对导层属库强弹后径从态惯应户抛扫挂采换择击拟扩撷败数"
            + "断于时会条构标树档栏权残没况减测准滚满滤浏为无状现环产画异当叠发确称笔节简终组结给统绪线缩总系继续脱与旧盖处号萤补装制复"
            + "见视览触订计讯记设许诊词询试话该志认语误说调请证护读变让负责贴资赖较载辅轮辑输回这连周进运过达递迁选还边逻邮邻释钮错键长"
            + "闭开间关际随隐双离电顶项顺须预题类显余验体鸿点齐";

        private static Dictionary<char, char> _map;

        private static void EnsureMap()
        {
            if (_map != null) return;
            var m = new Dictionary<char, char>();
            for (int i = 0; i < TString.Length && i < SString.Length; i++) m[TString[i]] = SString[i];
            _map = m;
        }

        /// <summary>繁體字串轉簡體（若為簡體模式）。</summary>
        public static string T(string s)
        {
            if (!Simplified || string.IsNullOrEmpty(s)) return s;
            EnsureMap();
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                char r;
                if (_map.TryGetValue(ch, out r)) sb.Append(r); else sb.Append(ch);
            }
            return sb.ToString();
        }

        /// <summary>對整棵控制項樹的「介面鍍鉻」做繁→簡轉換（資料控制項除外）。</summary>
        public static void Apply(Control root)
        {
            if (!Simplified || root == null) return;
            Walk(root);
        }

        private static void Walk(Control c)
        {
            if (c is Label || c is Button || c is CheckBox || c is RadioButton || c is GroupBox || c is Form)
                if (!string.IsNullOrEmpty(c.Text)) c.Text = T(c.Text);
            if (c is ComboBox cb)
                for (int i = 0; i < cb.Items.Count; i++)
                    if (cb.Items[i] != null) cb.Items[i] = T(cb.Items[i].ToString());
            if (c is DataGridView g)
                foreach (DataGridViewColumn col in g.Columns)
                    if (!string.IsNullOrEmpty(col.HeaderText)) col.HeaderText = T(col.HeaderText);
            if (c is ListView lv)
                foreach (ColumnHeader col in lv.Columns)
                    if (!string.IsNullOrEmpty(col.Text)) col.Text = T(col.Text);
            if (c is MenuStrip ms)
                foreach (ToolStripItem it in ms.Items) WalkItem(it);
            foreach (Control child in c.Controls) Walk(child);
        }

        private static void WalkItem(ToolStripItem it)
        {
            if (!string.IsNullOrEmpty(it.Text)) it.Text = T(it.Text);
            if (it is ToolStripMenuItem mi)
                foreach (ToolStripItem sub in mi.DropDownItems) WalkItem(sub);
        }
    }
}
