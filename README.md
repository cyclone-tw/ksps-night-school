# 國姓國小進修部管理系統

南投縣國姓國民小學進修部（補校）的教學管理系統，提供教學日誌填報、出缺席管理、薪資報表生成等功能。

## 系統架構

- **前端**：單頁式 HTML/CSS/JS 應用（GitHub Pages）
- **後端**：Google Apps Script (GAS) Web App
- **資料庫**：Google Sheets（8 個分頁）
- **報表**：GAS 伺服器端生成 XLS，自動存入指定 Google Drive 資料夾

## 功能

### 教師模式
- 教學日誌填報（含語音輸入、快速片語）
- 學生出缺席記錄（出席 / 請假切換）
- 同日期重複提交會覆蓋舊資料（提交前確認）

### 管理者模式（需密碼登入）
- 七日概覽儀表板（出席率、出席/請假/缺席名單）
- Google Sheets 原始資料快速連結
- 報表下載中心：
  - 教學日誌 XLS（A4 橫印格式）
  - 月薪資總表 XLS
  - 個人薪資條 XLS
  - 出缺席報表 XLS（自訂日期區間 + 學生選擇）
- 成績管理（開發中）

## Google Sheet 分頁結構

| 分頁 | 用途 |
|------|------|
| 系統設定 | 學校名稱、密碼、鐘點費、報表資料夾ID 等 |
| 人員名冊 | 教職員姓名、職稱、狀態、額外費用 |
| 學生名冊 | 學生姓名、狀態（在學/休學/輟學） |
| 課程設定 | 課程名稱、星期、授課教師 |
| 教學日誌 | 每日教學紀錄 |
| 出缺席記錄 | 每日出缺席狀態 |
| 成績設定 | 成績項目設定（預留） |
| 成績記錄 | 學生成績資料（預留） |

## 部署

### 前端（GitHub Pages）

```
https://cyclone-tw.github.io/ksps-night-school/
```

### 後端（GAS Web App）

```
https://script.google.com/macros/s/AKfycbxxaTfxJlZmqNBXc2gvTBb0rnUQpShm30Y8YFKfpHjIb8S5RLlrwzQz1xOIDLxf0W9j/exec
```

### 部署流程

```bash
# 推送 GAS 程式碼
cd gas
npx @google/clasp push --force

# 如果沒有新增 API scope，可用 clasp 部署
npx @google/clasp deploy -i AKfycbxxaTfxJlZmqNBXc2gvTBb0rnUQpShm30Y8YFKfpHjIb8S5RLlrwzQz1xOIDLxf0W9j

# 如果有新增 scope，必須在 GAS 編輯器手動部署：
# 部署 → 管理部署項目 → 編輯 → 新版本 → 部署
```

### 初次設定

1. 在 GAS 編輯器執行 `setupMoveTrigger()` — 設定每分鐘觸發器，自動移動報表到指定資料夾
2. 在 Google Sheet「系統設定」填入「報表資料夾ID」

## 技術備註

### GAS Web App 限制

- **不支援 POST 請求**：獨立部署的 GAS Web App 無法處理 POST，所有資料提交改用 GET + URL 參數
- **匿名模式無法使用 Drive API**：以 `ANYONE_ANONYMOUS` + `USER_DEPLOYING` 部署時，`DriveApp`、`UrlFetchApp`、`Advanced Drive Service` 均不可用。報表檔案移動改用時間觸發器（以 owner 權限執行）
- **clasp deploy 不觸發 OAuth 授權**：新增 scope 時必須在 GAS 編輯器手動部署

詳細開發紀錄見 [`docs/development-log.md`](docs/development-log.md)

## 專案結構

```
.
├── index.html              # 前端主頁（SPA）
├── favicon.png             # 網站圖示
├── manual.html             # 教師使用手冊
├── gas/
│   ├── Code.gs             # GAS 後端程式碼
│   ├── appsscript.json     # GAS 專案設定
│   └── .clasp.json         # clasp 設定（gitignore）
├── docs/
│   ├── development-log.md  # 開發日誌
│   └── plans/
│       ├── 2026-02-23-system-redesign-design.md
│       └── 2026-02-23-system-redesign-implementation.md
└── README.md
```
