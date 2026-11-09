# 詮栩 CNC 雲端系統

單人使用的 CNC 雲端工具入口。首頁提供「加工數量登記、加工數量查詢、庫存管理系統」三個功能；加工時間計算已整合到加工數量登記頁，不再使用獨立入口。

前端使用 React、TypeScript 與 Vite，雲端資料使用 Supabase Free，可由 GitHub Pages 自動部署。雲端模式僅允許指定的 Supabase 帳號登入；刪除、完整備份與還原會再要求一次帳號密碼。專案不含 Lovable 套件、執行環境或任何平台綁定。

## 已完成功能

- 深色細格線介面、桌面側欄與手機版選單
- 系統首頁：三個功能入口；只有進入「庫存管理系統」後才顯示庫存功能選單
- 可在登記頁手動新增或刪除產品編號與機台；變更會保存到雲端共用選單
- 刪除選單項目只會停止新登記使用，既有加工紀錄仍會保留並可從查詢頁查看
- 加工數量登記：日期、產品編號、機台、工時、數量與備註；產品編號和機台只能從已建立的下拉選單選擇
- 登記頁以彈出視窗提供工時相減，例如 `11 小時 5 分 - 3 小時 10 分 = 7 小時 55 分`，並可把結果帶回實際工時
- 加工數量查詢：依產品編號執行日期區間、最後日期或全部紀錄查詢；「查詢最後一筆」會列出該產品最後日期的所有機台並加總數量
- 三種加工查詢結果都可逐筆刪除；刪除後列表與統計會立即重算
- 加工統計：總天數、總數量、總時間、每日平均顆數與每顆平均秒數
- 訂單完工預估改為低頻工具彈出視窗；輸入訂單總數量後，以產品歷史每日平均產量計算已完成、剩餘數量、預估天數與完工日期
- 可在查詢頁永久刪除所選產品編號；加工紀錄與「管理產品編號」中的選項會在同一交易內一起刪除
- 只有「查詢全部」結果可下載為真正的 `.xlsx` 活頁簿；工作表、欄位、欄寬與底部統計列依提供的範例格式製作
- 登記頁可刪除最後一筆加工紀錄
- 庫存總覽：庫存總數、估算價值、本月入庫/出庫、近 14 天圖表
- 庫存盤點：直接輸入每項刀具的現場實際數量，不要求單價；系統保留原單價並以差異建立盤點調整紀錄
- 盤點頁可新增沒有價格的新刀具，也能搜尋、分類篩選並逐項更新數量
- 刀具資料：搜尋、四種分類篩選、新增、編輯、刪除
- 新增商品先選「刀片、鑽銑刀、刀柄、其他」，表單依分類顯示需要的欄位
- 規格、尺寸、斷屑槽、材質、廠牌、供應商與單位都是可輸入的下拉欄位；選項由現有刀具資料自動整理與去重
- 刀片品名依「規格－尺寸－斷屑槽－材質」組合；鑽銑刀依「規格－尺寸－材質」組合；刀柄與其他直接使用規格
- 商品編號依分類獨立編排：刀片 `T0001`、鑽銑刀 `M0001`、刀柄 `H0001`、其他 `O0001`
- 新增商品時可同時建立第一筆入庫
- 新增商品及入庫表單會依數量與單價即時計算採購總價
- 商品列可直接入庫、出庫，並顯示最新單價相對前次的漲跌百分比
- 歷史價格視窗：最新、最低、最高，以及每次入庫的日期、單價、價差、數量和供應商
- 出入庫商品可輸入編號、品名、廠牌、供應商或分類即時搜尋
- 一般入出庫的商品欄預設留空；從商品列快捷操作時才會自動選中該商品
- 入庫供應商會帶入商品目前的供應商，並支援從所有既有供應商下拉選擇或輸入新資料
- 每筆出入庫紀錄均可安全刪除；系統會重算後續餘額、商品庫存和最新單價，並拒絕會造成歷史負庫存的刪除
- 出入庫歷史：全文搜尋、類型與日期區間篩選
- 入庫與出庫彈窗、負庫存提示、重複送出防護
- 自訂登入名稱 `kuanlin`，實際身分由 Supabase Auth Email／密碼安全驗證
- RLS 與安全交易 RPC 固定綁定單一 Supabase UID，匿名訪客無法讀取或修改資料
- 登入狀態只保留在目前瀏覽器分頁；重新整理不會中斷，關閉分頁或瀏覽器後需要重新登入
- 原子化 PostgreSQL RPC：鎖定刀具資料列，同一交易更新庫存與歷史
- 首頁可由管理員將完整 Supabase 資料匯出為 Excel，整理後經格式檢查再整批還原
- 無 Supabase 設定時自動使用可操作的瀏覽器示範資料
- GitHub Actions / Pages 部署設定

## 先在本機執行

需要 Node.js 20 以上版本（建議 Node.js 22）。

```bash
npm install
npm run dev
```

終端機會顯示本機網址。直接開啟即可進入示範模式；示範資料只存在該瀏覽器的 `localStorage`。

正式建置：

```bash
npm run build
npm run preview
```

## 連接 Supabase Free

### 1. 建立資料表、RLS 與安全交易函式

1. 在 [Supabase](https://supabase.com/) 建立專案。
2. 進入專案的 SQL Editor。
3. 開啟 [`supabase/schema.sql`](supabase/schema.sql)，複製整份內容並執行。

`schema.sql` 是建立空白專案基礎資料表使用。已經升級為單一帳號安全模式後，不要再次執行舊的 `schema.sql` 或 `admin_data.sql`，以免恢復舊共用政策。

這份 SQL 會建立：

- `products` 刀具表
- `stock_movements` 不可由前端直接竄改的出入庫歷史
- `production_records` 加工數量、機台、工時與備註紀錄
- `production_options` 產品編號與機台下拉選單資料
- 入庫紀錄同時作為刀具的歷史價格來源，不需另一張價格表
- 基礎 Row Level Security policies（最後會由單一帳號安全升級檔取代）
- `create_product_with_initial_stock(...)` 安全建立刀具 RPC
- `register_stock_movement(...)` 安全出入庫 RPC
- `delete_stock_movement(...)` 安全刪除與完整庫存軌跡重算 RPC
- `set_stocktake_quantity(...)` 鎖定刀具並以實際盤點數量建立差異紀錄的 RPC；不要求或覆寫單價
- `delete_production_product_data(...)` 原子化永久刪除加工紀錄與產品編號選項 RPC

`register_stock_movement` 會鎖定刀具資料列、檢查共用資料範圍與負庫存，並以 `request_key` 防止同一次請求重送造成重複加減。

### 2. 新專案先啟用完整 Excel 備份／還原

只有全新空白 Supabase 專案才需要先執行 [`supabase/admin_data.sql`](supabase/admin_data.sql)。不要再設定舊的 8 位數管理密碼；下一步安全升級會移除舊密碼表，改用 Supabase 帳號密碼重新驗證。

首頁的「資料備份／還原」可以：

- 匯出刀具資料、出入庫歷史、加工數量與管理選項到同一份 Excel。
- 選擇整理後的 Excel，先檢查欄位、日期、數字、重複編號與資料筆數。
- 輸入「覆蓋資料」並再次通過帳號登入密碼後，以單一資料庫交易完整還原；發生錯誤會全部回復。

匯入會取代目前的雲端資料，務必保留一份未修改的原始備份。

### 3. 執行單一帳號安全升級

1. 先從舊網站下載一份完整 Excel 備份。
2. 確認 Supabase Authentication → Users 中已存在 `kuanlin@chyuansheu.com`，UID 為 `5568706d-caaa-459b-b842-13cdbc3929fc`。
3. 在 SQL Editor 建立 **New query**，完整貼上 [`supabase/secure_single_user.sql`](supabase/secure_single_user.sql) 後按 **Run**。
4. 成功畫面會顯示 `Success. No rows returned`。此檔會核對 UID 與 Email、把既有共用資料歸戶、關閉匿名權限並重建安全 RPC。

安全升級完成後，不要再執行 `schema.sql`、`admin_data.sql` 或舊的 `set_admin_password` 指令。

### 4. 填入前端環境變數

從 Supabase 專案設定的 API 頁面取得 Project URL 與 anon/public key。複製範例檔：

```bash
cp .env.example .env.local
```

Windows PowerShell 可使用：

```powershell
Copy-Item .env.example .env.local
```

編輯 `.env.local`：

```dotenv
VITE_SUPABASE_URL=https://你的專案.supabase.co
VITE_SUPABASE_ANON_KEY=你的_anon_key
```

重新執行 `npm run dev`。雲端模式會先顯示登入畫面；登入名稱為 `kuanlin`。

> `anon`／publishable key 本來就是供瀏覽器使用；絕對不要把 `service_role` key 放進 Vite、GitHub Pages 或任何 `VITE_` 變數。真正的資料保護由 Supabase Auth、RLS 與安全 RPC 負責。

首次使用或忘記密碼時，在登入畫面按「忘記密碼或尚未設定密碼」。請先在 Supabase Authentication → URL Configuration 將正式 GitHub Pages 網址加入 Redirect URLs。郵件連結回到網站後會顯示設定新密碼畫面；密碼限制為至少 4 位數字，且已公開使用過的舊管理密碼 `16866319` 不可再使用。

Supabase 專案本身也有密碼最短長度設定。請在 Auth 的 Password Security 設定中把 minimum password length 改成 `4`；否則前端雖接受 4 位數字，Supabase 仍會依雲端設定拒絕儲存。

## 部署到 GitHub Pages

專案已包含 [`.github/workflows/deploy.yml`](.github/workflows/deploy.yml)，推送到 `main` 後會自動建置與部署。

### 1. 建立 GitHub Repository 並推送

```bash
git init
git add .
git commit -m "Initial blade inventory app"
git branch -M main
git remote add origin https://github.com/你的帳號/你的專案.git
git push -u origin main
```

### 2. 新增 GitHub Actions secrets

在 Repository 的 **Settings → Secrets and variables → Actions** 新增：

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`

### 3. 啟用 Pages

在 **Settings → Pages → Build and deployment**，將 Source 設為 **GitHub Actions**。接著重新執行 Deploy to GitHub Pages workflow，或再推送一次 `main`。

Vite 已使用相對資源路徑，因此同時支援 `帳號.github.io` 與 `帳號.github.io/專案名稱/`，不需要手動修改 repository 名稱。

## 資料與安全設計

- 匿名角色沒有任何資料表或業務 RPC 權限。
- 每筆資料的 `user_id` 都必須等於指定 UID；RLS 同時檢查登入 JWT 與資料擁有者。
- 加工紀錄新增與查詢只開放給指定帳號；刪除前由 Supabase Auth 再次驗證帳號密碼。
- 前端可編輯刀具說明欄位，但沒有直接修改 `stock_quantity`、`unit_price` 或歷史資料的權限。
- 所有庫存增減只允許呼叫 `register_stock_movement`。
- 刪除出入庫只允許呼叫已登入帳號專用的 RPC，不能由瀏覽器直接刪除歷史表。
- RPC 從登入 JWT 取得 UID，不接受前端傳入任意 `user_id`。
- 刀具編號會依「刀片、鑽銑刀、刀柄、其他」四種分類分別安全遞增。
- 刪除刀具會一併刪除該刀具的歷史，且必須再次通過帳號密碼驗證。

## 專案結構

```text
blade-inventory/
├─ .github/workflows/deploy.yml  # GitHub Pages
├─ public/                       # 靜態圖示
├─ src/
│  ├─ lib/repository.ts          # Supabase / 示範資料存取
│  ├─ lib/supabase.ts            # Supabase client
│  ├─ lib/auth.ts                # 指定帳號與重新驗證
│  ├─ App.tsx                    # 頁面與互動
│  ├─ styles.css                 # 完整視覺與響應式樣式
│  └─ types.ts
├─ supabase/schema.sql           # Schema、RLS、日常功能 RPC
├─ supabase/admin_data.sql       # 管理員 Excel 備份／還原 RPC
├─ supabase/secure_single_user.sql # 單一帳號歸戶、RLS 與安全 RPC 升級
├─ .env.example
└─ vite.config.ts
```

## 常見問題

### 網站仍顯示「示範模式」

確認兩個環境變數都有正確值，且修改後已重新啟動開發伺服器或重新執行 GitHub Actions。

### 進入庫存管理後顯示資料庫錯誤

通常是還沒完整執行 `supabase/schema.sql`。回到 SQL Editor 重新執行整份檔案，再重新整理網站。

### 出庫顯示庫存不足

這是資料庫交易函式的必要防護。請先確認商品目前庫存，或先登記入庫。

### GitHub Pages 打開後空白

在 Actions 查看 build 是否成功，並確認 Pages 的 Source 是 GitHub Actions。若 workflow 沒有讀到兩個 secrets，網站仍應顯示示範模式，不會因此完全空白。

## 授權與備份

這是純前端加 Supabase 的個人專案。正式使用後，建議定期從 Supabase 匯出 `products` 與 `stock_movements`；資料庫層才是庫存的唯一正確來源。
