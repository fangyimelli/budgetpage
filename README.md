# 個人預算查詢網站（Google Apps Script）

## 專案檔案
- `Code.gs`：後端驗證、讀取 Google Sheet、寄送每月預算通知。
- `Index.html`：主頁面骨架。
- `Styles.html`：樣式。
- `JavaScript.html`：前端互動。
- `appsscript.json`：Apps Script 設定。

## Google Sheet 範例資料格式

### 1. users 工作表
| user_id | project_name | account | password | name | email |
| --- | --- | --- | --- | --- | --- |
| U001 | 教育部計畫 A | amy01 | pass1234 | 王小美 | amy@example.com |
| U002 | 國科會計畫 B | ben02 | pass5678 | 陳大文 | ben@example.com |

### 2. budget_summary 工作表
| user_id | project_name | 業務費 | 實支+未核銷 | 餘額 | 支用率 |
| --- | --- | --- | --- | --- | --- |
| U001 | 教育部計畫 A | 120000 | 48000 | 72000 | 40% |
| U002 | 國科會計畫 B | 80000 | 30000 | 50000 | 37.5% |

### 3. 個人分頁（例如 `U001`）
| 實支 | 日期 | 傳票 | 類別 | 購案編號 | 金額 | 說明 |
| --- | --- | --- | --- | --- | --- | --- |
| 是 | 2026/03/20 | V1140001 | 耗材 | P-001 | 3500 | 文具與會議資料印製 |
| 是 | 2026/03/10 |  | 差旅 | P-002 | 1200 | 高雄場次交通與住宿補助 |

## 部署成 Web App 步驟
1. 在 Google Drive 建立試算表，依照上方格式建立 `users`、`budget_summary` 與各使用者個人分頁。
2. 開啟 Apps Script 專案，將本專案檔案內容分別貼到對應檔案。
3. 確認 Apps Script 專案綁定到該 Google Sheet，或在同一份試算表中開啟「擴充功能 → Apps Script」。
4. 儲存後執行一次 `sendMonthlyBudgetEmails()` 或登入流程相關函式，完成權限授權。
5. 點選右上角「部署 → 新增部署」。
6. 類型選擇「網頁應用程式（Web app）」。
7. 執行身分建議選擇「我」。
8. 存取權限依需求設定；若僅內部使用，可選擇組織內使用者。
9. 部署完成後取得 Web App 網址，提供給使用者登入。

## 每月自動寄信 Trigger 設定方式
1. 在 Apps Script 左側選單點選「觸發條件」。
2. 點選右下角「新增觸發條件」。
3. 選擇要執行的函式：`sendMonthlyBudgetEmails`。
4. 部署版本選擇「Head」。
5. 事件來源選擇「時間驅動」。
6. 時間型態選擇「月計時器」。
7. 再指定每月日期與時段，例如每月 1 日上午 9 點。
8. 儲存後完成授權，之後系統就會每月自動寄送預算通知。
