# 開發日誌 — 國姓國小進修部管理系統

## 2026-02-23 系統重建

### 完成項目
- **Task 1-13**：依照 `docs/plans/2026-02-23-system-redesign-implementation.md` 完成全部實作
- GAS 後端：公開 API、管理者 API、四種報表生成（教學日誌、薪資總表、薪資條、出缺席報表）
- 前端 SPA：教師模式（教學日誌 + 出缺席填報）、管理者模式（總覽 + 報表下載 + 成績管理預留）
- GitHub Pages 部署：https://cyclone-tw.github.io/ksps-night-school/
- Google Sheet 初始化 API（`?action=init`）自動建立 8 個分頁和初始資料

### 技術決策
- **獨立 GAS 專案**（非容器綁定），使用 `SpreadsheetApp.openById()` 存取 Sheet
- **clasp** 管理 GAS 程式碼部署（`npx @google/clasp push --force` + `clasp deploy`）
- **Script ID**：`1qop1QNqQViUPjwCgxRoOJUssIuZFeYCDknfmUQiSfaTEiUpI4sCbHlCE`
- **Deployment ID**：`AKfycbxxaTfxJlZmqNBXc2gvTBb0rnUQpShm30Y8YFKfpHjIb8S5RLlrwzQz1xOIDLxf0W9j`
- **Sheet ID**：`1eaSKqrp7iQyW2yahpSV0a3ZT4A3jVo_lZjHpQQvcfNw`

---

## 2026-02-23 ~ 02-24 Bug 修復與功能改善

### POST 改 GET（v4）
- **問題**：獨立部署的 GAS Web App 不支援 POST 請求，同仁無法提交教學日誌和出缺席
- **解法**：將 `submit_log` 和 `submit_attendance` 改為 GET 請求，資料用 `encodeURIComponent(JSON.stringify(data))` 放在 URL 參數

### 同日期覆蓋（v5）
- **問題**：同一天重複提交會產生多筆資料，影響報表和薪資計算
- **解法**：`submitLog` 和 `submitAttendance` 改為同日期有舊資料就覆蓋（`setValues`），沒有就新增（`appendRow`）

### 管理者 Tab 顯示修復
- **問題**：CSS `.admin-only { display: none }` 優先級高於 JS inline style
- **解法**：改用 class-based 方式，在 `#navTabs` 加上 `admin-mode` class，用 `.nav-tabs.admin-mode .nav-tab.admin-only { display: block }` 覆蓋

### 七日概覽 Dashboard（v6）
- 「今日概覽」改為「七日概覽」，下拉選單可選近 7 天日期
- 出席率分母改用學生名冊「在學」人數
- 新增出席 / 請假 / 缺席學生名單（三欄顯示）
- 未標記的在學學生自動視為缺席

### 提交前確認覆蓋（v8）
- 新增 `check_date` API，提交前檢查該日期是否已有紀錄
- 有舊資料時跳出 `confirm()` 詢問是否覆蓋

### 報表合併儲存格修復（v9）
- **問題**：`ws.merge(ws.getRange(...))` 語法錯誤，Sheet 物件沒有 `merge()` 方法
- **解法**：改為 `ws.getRange(...).merge()`，修正 12 處

### Dashboard 在學學生過濾 + 自動重載 + Sheets 連結（v9~）
- 出缺席統計過濾掉休學 / 輟學學生
- 切到總覽頁籤自動重新載入最新資料
- 加上「開啟 Google Sheets 原始資料」超連結

### Favicon
- 加入 KSPS 進修部 logo 作為 `favicon.png`

---

## 2026-02-24 報表資料夾功能（歷經 26 版部署）

### 問題
希望生成的報表自動移到指定的 Google Drive 資料夾（用資料夾 ID 追蹤，搬移不受影響）。

### 嘗試過程與踩坑紀錄

> **這段紀錄特別重要，記錄了 GAS Web App 匿名模式下的 Drive API 限制。**

1. **DriveApp 直接操作**（`DriveApp.getFileById()` / `folder.addFile()` / `file.moveTo()`）
   - 全部失敗，錯誤：「你沒有呼叫 DriveApp.xxx 的權限」
   - 原因：Web App 以 `ANYONE_ANONYMOUS` + `USER_DEPLOYING` 部署時，匿名執行環境**無法使用 DriveApp**

2. **UrlFetchApp + Drive REST API**
   - 用 `ScriptApp.getOAuthToken()` 取 token，呼叫 Drive v3 PATCH API
   - 失敗，錯誤：「無法呼叫 UrlFetchApp.fetch，需要 script.external_request 權限」
   - 即使在 `appsscript.json` 加上 scope 也無效

3. **Advanced Drive Service**（`Drive.Files.update()`）
   - 在 `appsscript.json` 啟用 `enabledAdvancedServices`
   - 失敗，錯誤：「你沒有呼叫 drive.files.update 的權限」

4. **明確設定 `oauthScopes`**
   - 嘗試加入 `drive`、`drive.file`、`script.external_request` 等 scope
   - 在編輯器手動執行函式、授權、重新部署
   - 全部無效 — **Web App 匿名執行環境不繼承編輯器的 OAuth 授權**

### 根本原因

**GAS Web App 以 `ANYONE_ANONYMOUS` 存取時，匿名執行環境只有有限的 API 存取權限。`SpreadsheetApp` 可以用（因為是核心服務），但 `DriveApp`、`UrlFetchApp`、`Advanced Drive Service` 等需要額外 OAuth scope 的服務全部不可用。** 這是 Google Apps Script 的架構限制，不是授權問題。

### 最終解決方案：佇列 + 時間觸發器

- **報表生成時**：將 fileId 存入 `ScriptProperties` 佇列（`queueFileMove()`）
- **每分鐘觸發器**：`processMoveQueue()` 以 owner 權限執行，用 `DriveApp.getFileById().moveTo()` 移動檔案
- **一次性設定**：在 GAS 編輯器執行 `setupMoveTrigger()` 建立觸發器
- **資料夾 ID**：預設 `***REDACTED_FOLDER_ID***`，可在系統設定的「報表資料夾ID」欄位自訂

### 部署注意事項

- **clasp deploy 不會觸發 OAuth 重新授權**。新增 scope 時必須在 GAS 編輯器手動部署
- 每次 clasp push 後，需在編輯器「管理部署項目」→ 鉛筆 → 新版本 → 部署（如有新 scope 的話）
- 觸發器只需設定一次（`setupMoveTrigger()`），之後不用再動
