# 運維：分享權限與匿名填報穩定度

## 匿名填報不變

教學日誌／出缺席 **不需 Google 登入**。  
GAS Web App 部署為 `ANYONE_ANONYMOUS` + `USER_DEPLOYING`，以後端部署者身分寫入 Sheet。

管理者報表／七日總覽仍需管理者密碼。

## Drive 分享（建議）

下列檔案目前常只屬於 `cyclonetw@gmail.com`。匿名 API 仍可寫入，但學校帳號點開連結會「沒有權限」，容易被誤認為上傳失敗：

1. Google Sheet：`進修部GoogleSheet`
2. 報表資料夾：`進修部系統生成表件`（系統設定「報表資料夾ID」）

建議用部署者 Gmail 開啟 → 共用 → 加入 `ksps.ntct.edu.tw` 網域（檢視者即可；需要共編再給編輯）。

## 多帳號 404

Google 官方限制：瀏覽器同時登入多個 Google 帳號時，Apps Script Web App 可能回 404。  
請改用無痕視窗，或只保留一個登入帳號後重試。前端已會顯示此提示並自動重試。

## 報表搬移觸發器

報表建立後靠 `processMoveQueue`（每分鐘）搬到報表資料夾。  
若新檔卡在「我的雲端硬碟」根目錄，到 GAS 編輯器執行一次 `setupMoveTrigger()`。

## 部署 checklist

```bash
cd gas
pnpm dlx @google/clasp@2.4.2 push --force
pnpm dlx @google/clasp@2.4.2 deploy -i AKfycbxxaTfxJlZmqNBXc2gvTBb0rnUQpShm30Y8YFKfpHjIb8S5RLlrwzQz1xOIDLxf0W9j -d "stabilize"
```

若新增 OAuth scope，必須在 GAS 編輯器手動「管理部署項目」產生新版本。  
前端 `API_URL` 的 deployment id 變更時需同步 `index.html`。
