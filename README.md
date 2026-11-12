# 詮栩 CNC 雲端系統

單人使用的 CNC 雲端工具入口。首頁提供「加工管理系統、庫存管理系統、資料備份／還原」三個功能；加工管理系統包含總覽、加工數量登記、加工數量查詢、訂單完工預估、基礎資料設定與資料匯出／刪除分頁。

前端使用 React、TypeScript 與 Vite，雲端資料使用 Supabase Free，可由 GitHub Pages 自動部署。管理員使用指定的 Supabase 帳號登入；訪客可免登入唯讀查看資料，但無法新增、修改、刪除、盤點、匯出或還原。刪除、完整備份與還原會再要求一次管理員帳號密碼。專案不含 Lovable 套件、執行環境或任何平台綁定。

## 已完成功能

- 深色細格線介面、桌面側欄與手機版選單
- 系統首頁：加工管理系統、庫存管理系統與資料備份／還原三個功能入口
- 加工管理系統使用獨立側邊分頁；基礎資料設定可新增或刪除產品編號與機台，也可編輯產品編號並同步取代全部既有加工紀錄中的舊編號
- 刪除選單項目只會停止新登記使用，既有加工紀錄仍會保留並可從查詢頁查看
- 加工數量登記：日期、產品編號、機台、工時、數量與備註；產品編號和機台只能從已建立的下拉選單選擇
- 登記頁以彈出視窗提供工時相減，例如 `11 小時 5 分 - 3 小時 10 分 = 7 小時 55 分`，並可把結果帶回實際工時
- 加工數量查詢：依產品編號執行日期區間、最後日期或全部紀錄查詢；「查詢最後一筆」會列出該產品最後日期的所有機台並加總數量
- 三種加工查詢結果都可逐筆刪除；刪除後列表與統計會立即重算
- 加工統計：總天數、總數量、總時間、每日平均顆數與每顆平均秒數
- 訂單完工預估使用獨立側邊分頁；以歷史「每機台每小時產量」乘上預計機台數及每日工時，計算剩餘工作天與完工日期
- 完工排程可選擇依歷史紀錄自動判斷週六／週日、週一至週六或純平日；國定假日與補假依人事總處 2026～2027 官方日曆排除
- 可輸入訂單交期，自動顯示能否在交期內完成，以及預計提前或延後的工作天數
- 訂單完工預估會在目前瀏覽器保存所選產品、訂單數量、交期、排程、機台數與每日工時；切換功能或重新整理後會自動還原（草稿不會跨手機與電腦同步）
- 預估結果分開顯示「加工日產能、排程平均日產量、實際加工天數、預估日曆天數」；切換純平日或週一至週六時，平均日產量、日曆天數及完工日期會立即重算
- 國定假日資料集中在 `src/lib/taiwanHolidays.ts`；官方公布新年度辦公日曆後需將新年度日期加入此檔案
- 資料匯出與刪除使用同一個側邊分頁；選擇產品編號後會直接顯示全部加工紀錄與統計，可選擇下載 Excel 或永久刪除
- Excel 下載會輸出真正的 `.xlsx` 活頁簿；工作表、欄位、欄寬與底部統計列依提供的範例格式製作
- 永久刪除會把加工紀錄與「管理產品編號」中的選項在同一交易內一起刪除，並再次要求管理員密碼
- 登記頁可刪除最後一筆加工紀錄
- 庫存總覽：庫存總數、估算價值、本月入庫/出庫、近 14 天圖表
- 庫存盤點：直接輸入每項刀具的現場實際數量，不要求單價；系統保留原單價並以差異建立盤點調整紀錄
- 盤點頁可新增沒有價格的新刀具，也能搜尋、分類篩選並逐項更新數量
- 刀具資料：搜尋、四種分類篩選、新增、編輯、刪除
- 手機條碼掃描：可使用 iPhone 相機或 XB-6208-RB 藍牙掃描器辨識一維／二維條碼；掃描中、成功、未綁定會顯示不同顏色與明確提示，成功時播放提示音並自動移到結果
- 新增或編輯刀具時可自動產生不重複的 13 位內部條碼，並立即顯示相同內容的 QR Code 預覽；仍可保留掃描或手動輸入方式
- 「掃描刀具條碼」視窗提供 XB-6208-RB 的 iOS 鍵盤顯示／隱藏設定 QR Code，不必另外尋找紙本說明書
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
- 出入庫歷史：可掃描或輸入刀具條碼搜尋，也可依編號、品名、廠牌、供應商、備註、類型與日期區間篩選
- 出入庫、庫存盤點與刀具資料的搜尋框支援連續掃描；掃描器送出 Enter 後會提示成功或失敗，1.8 秒後只清空可輸入文字並重新聚焦，上一筆搜尋結果會保留，下一次掃描會直接取代上一筆條碼
- 入庫與出庫彈窗、負庫存提示、重複送出防護
- 自訂登入名稱 `kuanlin`，實際身分由 Supabase Auth Email／密碼安全驗證
- 登入名稱欄位只接受英文字母與數字，並關閉手機鍵盤的自動大寫、拼字修正與拼字檢查
- 登入畫面提供「訪客查看」；訪客可讀取正式雲端資料，但所有寫入功能與資料匯出均停用
- RLS 與安全交易 RPC 固定綁定單一 Supabase UID；匿名訪客只有四張業務資料表的 SELECT 權限
- 登入狀態只保留在目前瀏覽器分頁；重新整理不會中斷，關閉分頁或瀏覽器後需要重新登入
- 原子化 PostgreSQL RPC：鎖定刀具資料列，同一交易更新庫存與歷史
- 首頁可由管理員將完整 Supabase 資料匯出為 Excel，整理後經格式檢查再整批還原；備份檔另含可直接列印的「QR Code 標籤」工作表
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
- 自動建立「QR Code 標籤」工作表，以兩欄標籤排版已綁定條碼的刀具；QR Code 內容與「刀具資料」的條碼文字相同。
- 選擇整理後的 Excel，先檢查欄位、日期、數字、重複編號與資料筆數。
- 輸入「覆蓋資料」並再次通過帳號登入密碼後，以單一資料庫交易完整還原；發生錯誤會全部回復。

匯入會取代目前的雲端資料，務必保留一份未修改的原始備份。

### 3. 執行單一帳號安全升級

1. 先從舊網站下載一份完整 Excel 備份。
2. 確認 Supabase Authentication → Users 中已存在 `kuanlin@chyuansheu.com`，UID 為 `5568706d-caaa-459b-b842-13cdbc3929fc`。
3. 在 SQL Editor 建立 **New query**，完整貼上 [`supabase/secure_single_user.sql`](supabase/secure_single_user.sql) 後按 **Run**。
4. 成功畫面會顯示 `Success. No rows returned`。此檔會核對 UID 與 Email、把既有共用資料歸戶、關閉匿名寫入權限並重建安全 RPC。

安全升級完成後，不要再執行 `schema.sql`、`admin_data.sql` 或舊的 `set_admin_password` 指令。

### 3.1 啟用訪客唯讀資料

在 SQL Editor 建立 **New query**，完整貼上 [`supabase/enable-guest-readonly.sql`](supabase/enable-guest-readonly.sql) 後按 **Run**。

這份 SQL 只授予 `anon` 角色以下權限：

- 讀取 `products`
- 讀取 `stock_movements`
- 讀取 `production_records`
- 讀取 `production_options`

訪客沒有 INSERT、UPDATE、DELETE 或 RPC 執行權限。由於訪客不需要身分驗證，任何取得網站網址或 Supabase 公開端點的人都可能讀到這四張資料表；請勿存放不適合公開查看的個人或機密內容。

### 3.2 啟用加工產品編號編輯

在 SQL Editor 建立 **New query**，完整貼上 [`supabase/enable-production-product-code-edit.sql`](supabase/enable-production-product-code-edit.sql) 後按 **Run**。

這份 SQL 會建立只有指定管理員帳號能執行的安全 RPC。從「加工管理系統 → 基礎資料設定 → 管理產品編號」修改編號時，會在同一筆資料庫交易內同步更新產品編號選項與全部既有加工紀錄；若新編號已存在則會拒絕修改，避免兩個產品的資料意外合併。

### 啟用刀具條碼掃描資料欄位

在 SQL Editor 建立 **New query**，完整貼上 [`supabase/enable-product-barcodes.sql`](supabase/enable-product-barcodes.sql) 後按 **Run**。

這份 SQL 會新增刀具條碼欄位、避免同一條碼重複綁定，並把條碼加入完整 Excel 備份與還原。網站部署後，在庫存總覽、出入庫、庫存盤點或刀具資料頁按「掃描條碼」即可使用。iPhone 必須從正式 HTTPS 網址開啟，並允許 Chrome 使用相機。

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

重新執行 `npm run dev`。雲端模式會先顯示登入畫面；管理員登入名稱為 `kuanlin`，訪客可直接按「訪客查看」。

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

- 匿名角色只能 SELECT 四張業務資料表，沒有 INSERT、UPDATE、DELETE 或業務 RPC 權限。
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
├─ public/                       # 公司 Logo、網站圖示與 iOS 鍵盤切換 QR Code
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
├─ supabase/enable-guest-readonly.sql # 免登入訪客唯讀 RLS 與資料表權限
├─ supabase/enable-product-barcodes.sql # 刀具條碼欄位、唯一性與 Excel 備份升級
├─ supabase/enable-production-product-code-edit.sql # 安全修改加工產品編號並同步歷史紀錄
├─ .env.example
└─ vite.config.ts
```

### 修改網站字體大小

開啟 `src/styles.css`，檔案最前方有「全站字體大小集中設定區」。所有固定的 `font-size` 與 `font:` 縮寫尺寸都已集中成獨立變數，並分成共用／電腦版、平板版、電腦版補充、手機版等區域。可在 VS Code 按 `Ctrl + F` 搜尋畫面中文名稱，例如「實際工時（時）」、「訂單總數量」或「庫存盤點」，直接修改對應變數的 `px` 數值。下方各功能樣式只引用變數，不必再從整份 CSS 尋找數字。

## 常見問題

### 網站仍顯示「示範模式」

確認兩個環境變數都有正確值，且修改後已重新啟動開發伺服器或重新執行 GitHub Actions。

### 進入庫存管理後顯示資料庫錯誤

通常是還沒完整執行 `supabase/schema.sql`。回到 SQL Editor 重新執行整份檔案，再重新整理網站。

### 出庫顯示庫存不足

這是資料庫交易函式的必要防護。請先確認商品目前庫存，或先登記入庫。

### GitHub Pages 打開後空白

在 Actions 查看 build 是否成功，並確認 Pages 的 Source 是 GitHub Actions。若 workflow 沒有讀到兩個 secrets，網站仍應顯示示範模式，不會因此完全空白。

### iPhone 按「開啟相機掃描」後沒有畫面

確認使用的是 `https://` 正式網址，並到 iPhone 的 Chrome／網站相機權限允許使用相機。若相機暫時無法使用，可先在掃描視窗的「電腦測試／相機無法辨識時」欄位手動輸入條碼，測試查詢、建立與綁定流程。

相機掃描預設使用後鏡頭、1080p、30fps，並在瀏覽器支援時啟用連續對焦與約 1.35 倍變焦。掃描時先讓手機與條碼保持約 20～30 公分，將條碼置於框線中央後再慢慢靠近；避免手機貼得太近或燈光直接反射條碼貼紙。

### iPhone 連接 XB-6208-RB 後，手機鍵盤不會顯示

這是 iPhone 將藍牙掃描器視為外接鍵盤的正常行為。進入庫存功能的「掃描刀具條碼」視窗，展開「XB-6208-RB：顯示／隱藏 iOS 鍵盤」，再用掃描器掃描畫面中的 QR Code，即可切換手機鍵盤。掃描後若鍵盤沒有立刻出現，請重新點一下要輸入的欄位。這個設定碼只調整掃描器與 iOS 鍵盤，不會變更網站資料。

## 授權與備份

這是純前端加 Supabase 的個人專案。正式使用後，建議定期從 Supabase 匯出 `products` 與 `stock_movements`；資料庫層才是庫存的唯一正確來源。
