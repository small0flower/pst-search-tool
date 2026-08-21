# Outlook PST 搜尋工具（PstSearchTool）

一個獨立於 Outlook 的 PST 郵件搜尋工具，用於解決 **Outlook 2007 在 Windows 10/11 上搜尋失效** 的已知問題：
Outlook 2007 的「即時搜尋」依賴舊版 Windows Desktop Search（WDS），而 Windows 10/11 已將其移除並換成不相容的新版 Windows Search，導致 Outlook 2007 的搜尋功能直接失效（[Microsoft Q&A 回報](https://learn.microsoft.com/en-us/answers/questions/4683283/updated-to-windows-11-now-in-outlook-2007-can-no-l)、[中文討論](https://im5481.com/2025/01/03/outlook-2007-2013%E5%9C%A8windows-11%E7%84%A1%E6%B3%95%E7%B4%A2%E5%BC%95%E6%90%9C%E5%B0%8B%E9%83%B5%E4%BB%B6/)）。微軟沒有提供修復，本工具以「外部索引 + 點擊於 Outlook 開啟原件」補足此缺口。

> ⚠️ 本工具**不會修改**你的 PST 檔案——只讀取郵件內容建立索引；點擊搜尋結果時，由 Outlook 開啟該封郵件的**原件**。

---

## 功能

- **兩種郵件來源**
  - **目前 PST**：列出現行 Outlook 設定檔內掛載的所有存放區（PST/OST）。
  - **外掛 PST**：瀏覽目錄（含子資料夾）挑選 .pst 檔，自動加入 Outlook 設定檔後建立索引（之後可於 Outlook「檔案 → 資料檔管理」移除）。
- **資料夾選擇**：樹狀勾選要索引的資料夾（勾選資料夾會包含其子資料夾）；快速按鈕：收件匣／寄件匣／全選／清空。
- **中文全文索引**：自建倒排索引（bigram + trigram），**2 個字元以上的中文子字串即可精確搜尋**（如「鴻海」「王小明」「訂單編號」）；英文為不分大小寫的整詞比對。
- **搜尋過濾**：多關鍵字（空白分隔，AND 邏輯）、資料夾（全部／收件匣／寄件匣／自訂）、日期範圍、寄件者。
- **點擊開啟**：雙擊搜尋結果，由 Outlook 開啟該封郵件原件（使用 MAPI EntryID，非另存副本）。
- **索引管理**：建立／更新索引（增量：未變更的資料夾自動略過）、完整重建索引、取消；進度列與狀態顯示。
- **索引持久化**：索引存放於 `%APPDATA%\PstSearchTool\index.db`，關閉程式後仍可搜尋；搜尋不需 Outlook 開啟。

## 系統需求

| 項目 | 需求 |
|---|---|
| 作業系統 | Windows 7 SP1 / Windows 10 / Windows 11（x86 與 x64 皆可） |
| 執行環境 | .NET Framework 4.8（安裝程式會檢查；[下載](https://dotnet.microsoft.com/download/dotnet-framework/net48)） |
| Outlook | Outlook 2007 或更新版本（COM 互通，含 2010/2013/2016/2019/365） |
| 位元數 | 工具為 32 位元（x86）——因為 Outlook 2007 是 32 位元程式，COM 必須同 bitness |

## 快速開始

1. 執行安裝程式 `PstSearchTool-Setup-1.0.3.exe`（或直接解壓縮 `dist` 內的檔案執行 `PstSearchTool.exe`）。
2. **選擇來源**：預設「目前 Outlook 中的 PST」會列出 Outlook 設定檔內的存放區；或選「外掛 PST」→「瀏覽…」挑選 .pst 檔。
3. **勾選來源與資料夾**：在左側勾選存放區（預設勾選主要信箱），點選存放區後在資料夾樹勾選要索引的資料夾（預設勾選收件匣、寄件匣）。
4. 按 **「建立/更新索引」**，等待完成（首次建立大型 PST 可能需要一段時間；之後未變更的資料夾會自動略過，很快）。
5. 在上方輸入關鍵字（可多個、空白分隔），可加日期／寄件者／資料夾過濾，按 **「搜尋」**。
6. **雙擊結果列**即可在 Outlook 中開啟該封郵件。

## 搜尋語法與行為

- 多關鍵字以**空白**分隔，為 **AND** 邏輯（全部都要出現）。
- 中文：**2 字元以上**的子字串精確比對（例如「鴻海」找得到「鴻海出貨通知」）。
- 英文：整詞、不分大小寫（「order」找得到「Order」「ORDER」）。
- 單一字元查詢（如「王」）會自動改用較慢的全文掃描（仍正確，但大資料庫下會慢）。
- 結果預設依日期新→舊排序，最多顯示 500 筆。

## 索引技術說明

- **自建倒排索引**（`postings` 表：token → 郵件 id），**不依賴 SQLite FTS5 擴充**。原因：官方 System.Data.SQLite 二進位對 FTS5 的支援不一（[歷史回報](https://stackoverflow.com/questions/37565423/sqlite-no-such-module-fts5-error-with-system-data-sqlite-dll-1-0-101-0)、[官方論壇](https://sqlite.org/forum/forumpost/87420fddbe3494fa?t=c&unf)），自建索引保證任何 SQLite 版本皆可運作，且額外支援 2 字元中文搜尋。
- 分詞：中文連續段產生 2 字元 bigram 與 3 字元 trigram；英文小寫整詞。查詢時依關鍵字長度取對應 token 做交集，達到精確子字串比對。
- **增量更新**：每個資料夾記錄「郵件數 + PST 檔大小/修改時間」快照；未變更則略過，變更才重新索引該資料夾（含子資料夾）。注意：若某封舊信被「就地編輯」且郵件數與檔案都未變，可能不會被偵測到——需要時請用「完整重建索引」。
- **安全讀取**：所有位址欄位（寄件者/收件者等）透過 MAPI `PropertyAccessor` 讀取，避免 Outlook 的「程式嘗試存取電子郵件地址資訊」安全性警告；主旨／內文／時間等一般屬性不會觸發警告。

## 已知限制（v1）

- 附件內容（PDF/Word/Excel 內文）**不**索引，僅索引主旨／寄件者／收件者／內文。
- 首次索引大型 PST（數萬封）需較長時間，建議安排在非使用時段；之後增量更新很快。
- 外掛 PST 加入 Outlook 設定檔後會**持續掛載**（與你手動「新增資料檔」相同），可隨時在 Outlook 移除。
- Outlook 開啟中不建議以其他程式直接讀取該 PST 檔案；本工具一律透過 Outlook COM 存取，無此風險。
- 單一字元中文查詢較慢（見上）。

## 疑難排解

| 症狀 | 處理 |
|---|---|
| 「找不到 Outlook.Application」 | 未安裝 Outlook，或安裝了 64 位元 Outlook 而本工具誤以 64 位元執行（正常情況不會發生，x86 已固定）。 |
| Outlook 彈出「程式嘗試存取…」警告 | 點選「允許」（並可勾選允許數分鐘）；正常情況很少出現。 |
| 外掛 PST 掛載失敗 | 該檔可能已損毀或被其他程式獨佔；檢查路徑與檔案是否仍存在。 |
| 搜尋結果無法開啟 | 該郵件所屬的 PST 已從 Outlook 移除；請重新掛載（或重新建立索引）。 |
| 索引資料庫損壞 | 刪除 `%APPDATA%\PstSearchTool\index.db`（與 `-wal`/`-shm` 檔）後重新建立索引。 |

## 開發與建置

### 專案結構

```
pst-search-tool/
├── src/PstSearch/           # C# .NET Framework 4.8 WinForms（x86）
│   ├── Outlook/             # COM 服務層（晚期繫結，相容 Outlook 2007+）
│   ├── Data/IndexStore.cs   # SQLite 儲存層（倒排索引、快照、搜尋）
│   ├── Indexing/Indexer.cs  # 索引協調（Outlook 讀取 → SQLite 寫入）
│   ├── Search/              # Tokenizer（分詞）、Snippet（摘要）
│   ├── Config/              # 設定持久化（%APPDATA%\PstSearchTool\config.xml）
│   └── UI/                  # 主介面
├── build/                   # build.ps1（本機建置）、installer.iss（Inno Setup）
├── .github/workflows/       # GitHub Actions：Windows 建置 + 安裝程式
└── _check/                  # 僅供 Linux/mono 編譯檢查用的 stub（不屬專案）
```

### 本機建置（Windows）

需求：.NET SDK 8+（或 Visual Studio 2022 Build Tools）。

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1
# 輸出在 dist\；安裝程式：以 Inno Setup 6 的 ISCC.exe 編譯 build\installer.iss
```

### CI 建置（GitHub Actions）

推送到 GitHub（main 分支）或手動觸發 `build` workflow，即會自動在 `windows-latest` 建置並上傳成品（應用程式 + 安裝程式）。

## 授權

- 本工具為個人／內部使用工具，程式碼以 MIT 授權釋出（見 LICENSE）。
- 使用的第三方元件皆為免費／公眾領域：SQLite（public domain）、System.Data.SQLite（public domain）。
- v2 若導入 libpst（GPL）或 pypff（LGPL）做「直接讀檔」加速，需注意授權差異。

## Roadmap（v2 規劃）

- 直接讀檔模式（libpst/pypff）：Outlook 關閉時更快建立首次索引（不修改設定檔）。
- 附件內容索引（PDF/Word/Excel 文字抽取）。
- 搜尋結果匯出 CSV／另存 .msg。
- 關聯字（同客戶／同主旨串）與統計（依寄件者／日期分布）。
