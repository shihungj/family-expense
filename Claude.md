# 家庭每月支出明細系統 CLAUDE.md

## 專案架構
- 前端：HTML + CSS + Vanilla JS，單一檔案 index.html
- 後端：Google Apps Script（Code.gs）
- 資料庫：Google Sheets
- 部署：GitHub Pages（https://shihungj.github.io/family-expense/）

## Google Sheets 工作表結構

### Transactions
欄位順序：入帳日、消費日、銀行、消費明細、金額、歸屬、帳單月份、批次編號
- 日期格式：YYYY/MM/DD（純文字）
- 帳單月份格式：YYYY/MM（純文字）
- 金額：數字，讀取後必須 parseFloat() 處理
- 歸屬允許值：世鴻應付、慧鳳應付、共同支付、未分類

### BankSettings
欄位：銀行、結帳日、銀行縮寫
- 玉山銀行、13、ESUN
- 中國信託、25、CTBC
- 富邦銀行、26、FUBO
- 永豐銀行、26、SINO
- 聯邦銀行、12、UNIO

### MerchantRules
欄位：關鍵字、歸屬
- 新增規則插入第二列（標題列下方）

### SystemSettings
欄位：key、value
- 儲存帳號密碼，密碼以 Base64 編碼

### ImportLog（新增）
欄位：批次編號、銀行名稱、帳單月份、筆數、匯入時間
- 每次匯入成功後自動寫入，最新記錄插入第二列
- 匯入時間格式：YYYY/MM/DD HH:mm

## 批次編號規則
格式：銀行縮寫-YYYYMM，例如 ESUN-202605
匯入前檢查重複，重複則阻擋並提示

## KPI 計算規則
- 世鴻應付 ＝ 世鴻應付總額 ＋ 共同支付總額 ÷ 2（餘數歸世鴻無條件進位）
- 慧鳳應付 ＝ 慧鳳應付總額 ＋ 共同支付總額 ÷ 2（無條件捨去）
- 共同支付 ＝ 共同支付原始總額
- 本期總支出 ＝ 全部交易金額總和

## getKPI 回傳格式
```json
{
  "shihong": 0,
  "huifeng": 0,
  "common": 0,
  "total": 0,
  "shihongLastMonth": 0,
  "huifengLastMonth": 0,
  "commonLastMonth": 0,
  "totalLastMonth": 0,
  "transactionCount": 0,
  "importedBanks": [],
  "notImportedBanks": [],
  "paymentTools": {
    "linepay": { "count": 0, "amount": 0 },
    "icashpay": { "count": 0, "amount": 0 },
    "easycard": { "count": 0, "amount": 0 },
    "easywallet": { "count": 0, "amount": 0 },
    "carmoji": { "count": 0, "amount": 0 },
    "creditcard": { "count": 0, "amount": 0 }
  },
  "maxSingleAmount": 0,
  "topBankByAmount": "",
  "topBankByCount": ""
}
```

## getRecentImports 回傳格式
```json
{
  "records": [
    {
      "batchId": "ESUN-202605",
      "bankName": "玉山銀行",
      "billingMonth": "2026/05",
      "count": 18,
      "importedAt": "2026/05/15 14:32"
    }
  ]
}
```
- 從 ImportLog 工作表讀取，取前5筆
- totalLastMonth 累計上期所有交易金額（含未分類），與 shihongLastMonth + huifengLastMonth + commonLastMonth 加總可能略有差異，行為與本期一致

## 支付工具關鍵字（消費明細欄位，不分大小寫）
- LINE Pay：包含「連加」或「連支」
- iCASH Pay：包含「icash pay」
- 悠遊卡：包含「悠遊卡」
- 悠遊付：包含「悠遊付」
- 車麻吉：包含「中油條碼_Autopass」
- 信用卡：以上都不符合

## CSV 匯入格式
欄位：入帳日、消費日、銀行、消費明細、金額
- 第一列標題列固定跳過
- 空白列忽略
- 帳單月份由使用者手動選擇
- 可選範圍：當月往前推6個月，當月不可選
- 預設選上個月
- 選超過上個月跳橘色警示（可繼續）

## 安全機制
- 前端 sessionStorage 儲存登入 Token
- 每頁載入檢查 Token，不存在跳回登入頁
- 登入後統一進入首頁
- 密碼變更後跳回變更密碼頁顯示成功訊息
- Apps Script 每次請求需帶驗證碼：family-expense-2026

## 視覺風格
溫暖舒適家庭風格，字體 Noto Sans TC

## 配色規範
- 頁面最外層背景：#FAF7F2
- 主內容區背景：#FCFAF7
- 頂部列背景：#FAF7F2（與頁面同色，無邊框）
- 側邊欄底色：#FCFAF7，右側邊框 1px solid #E9E2DB
- 卡片背景：#FFFFFF
- 卡片邊框：#E9E2DB
- 卡片陰影：0 1px 3px rgba(0,0,0,0.06)
- 主題色（暖橘）：#FF8A3D
- 輔助色（暖綠）：#4CAF7D
- 輔助色（暖紫）：#A78BFA
- 輔助色（暖藍）：#60A5FA
- 文字主要：#1F2937
- 文字次要：#6B7280
- 文字輔助：#9CA3AF
- 支出上升：#EF4444
- 支出下降：#22C55E
- 邊框／分隔線：#E9E2DB
- 成功：#22C55E
- 警示：#F59E0B
- 錯誤：#EF4444
- 已匯入綠：#22C55E
- 未匯入橘：#FF8A3D

## 首頁版面結構
1. 左側固定選單（淺米色側邊欄 #FCFAF7）
   - Logo：logo.png，width 160px
   - 選單高亮：background #FFF5EE，文字 #FF8A3D，左側 3px solid #FF8A3D
   - 底部插圖：sofa.png，width 100%
2. 頂部列：帳單月份選擇器（左右箭頭＋點擊數字跳出年月選擇器）＋右側使用者名稱＋登出（無鈴鐺）
3. KPI 四張卡片（各含圖示、金額、較上期、筆數、最高單筆）
   - 世鴻消費金額：圖示底色 #FFF3E8，SVG 人形 #FF8A3D，金額 #FF8A3D
   - 慧鳳消費金額：圖示底色 #F3F0FF，SVG 人形 #A78BFA，金額 #A78BFA
   - 共同支出金額：圖示底色 #E8F8EF，SVG 雙人 #4CAF7D，金額 #4CAF7D
   - 家庭總支出：圖示底色 #FFF3E8，SVG 房子 #FF8A3D，金額 #FF8A3D
4. 本期 vs 上期對比（四欄卡片）
   - 各卡顯示：名稱、增減幅、本期金額、上期金額、差異
5. 統計參考（兩欄）：支付工具統計（含筆數與金額）＋ 消費行為參考
6. 帳單匯入進度（極簡一行：X / 5 已完成）

## 帳單匯入頁版面結構
1. 上方寬版卡片（左 45% : 右 55%，垂直漸變分隔線）
   - 左側：說明文字、帳單月份選擇器、匯入按鈕、注意事項
   - 右側：csv_input.png 配圖
2. 下方左右並排
   - 左：本月匯入狀態（五間銀行，超過高度出現捲動軸）
   - 右：最近匯入記錄表格（#、銀行、帳單月份、筆數、匯入時間，超過5筆出現捲動軸，斑馬紋）

## 全部交易明細頁規格
- 欄位：項次、帳單月份、銀行、消費日、消費明細、金額、歸屬、操作
- 篩選：銀行下拉、消費明細關鍵字、歸屬下拉
- 多選功能＋批次修改歸屬＋批次刪除
- 新增交易按鈕（右上角）
- 表格樣式：容器背景 #FFFFFF、表頭 #FBF7F2、列高 56px、斑馬紋、hover #FFF6EC
- 捲動軸：寬 8px，track #F3ECE5，thumb #D8CFC4

## 對帳單預覽頁規格（待實作）
- Tab 切換：慧鳳對帳單（預設）、世鴻對帳單
- 慧鳳：應付總金額＋慧鳳應付明細＋共同支付明細（含÷2捨去後金額）
- 世鴻：應付總金額＋世鴻應付明細＋共同支付明細（含÷2進位後金額）
- 匯出 PDF 使用 html2pdf.js
- PDF 檔名：慧鳳對帳單_202605.pdf / 世鴻對帳單_202605.pdf

## 視覺設計規範

### 按鈕樣式（全站統一）
1. 主要操作按鈕（實心底色，border-radius 8px）：
   - Primary：background #FF8A3D，color #FFFFFF
   - Secondary：background #F3F0FF，color #A78BFA
   - Danger：background #FEE2E2，color #EF4444
   - Ghost：background #F3F4F6，color #6B7280
2. 表格內操作按鈕（白底有色邊框，border-radius 999px）：
   - 修改：border #FF8A3D，color #FF8A3D
   - 刪除：border #EF4444，color #EF4444

### 歸屬標籤（全站統一，border-radius 999px）
- 世鴻應付：background #FFF3E8，color #FF8A3D，border #FFD4A8
- 慧鳳應付：background #F3F0FF，color #A78BFA，border #DDD6FE
- 共同支付：background #E8F8EF，color #4CAF7D，border #A7F3C4
- 未分類：background #F3F4F6，color #6B7280，border #E5E7EB
- 注意：歸屬標籤值（世鴻應付／慧鳳應付／共同支付／未分類）與後端對應，不可更改

### 純文字名稱（全站）
- 頁面標題、卡片名稱、表格欄位等純文字顯示：
  - 世鴻消費金額、慧鳳消費金額、共同支出金額

### 關鍵字標籤（分類規則管理頁）
- background #FFF3E8，color #FF8A3D，border #FFD4A8，border-radius 999px，帶 × 刪除符號

### 表格樣式（全站統一）
- 外框 border 1px solid #E9E2DB，border-radius 12px
- 標題列：background #FAF7F2，color #6B7280
- 資料列 hover：background #FFF8F3
- 內格線：border-bottom 1px solid #E9E2DB
- 金額欄位靠右對齊，font-weight 600

### 輸入框樣式（全站統一）
- border 1px solid #E9E2DB，border-radius 8px
- focus：border-color #FF8A3D，box-shadow 0 0 0 3px rgba(255,138,61,0.15)
- placeholder color #9CA3AF

### 卡片圓角與間距
- border-radius 12px
- 卡片間距 gap 16px～24px
- 主內容區左右 padding 1.5rem

## 登入頁
- Logo：login.png，width 180px，不可被 CSS 拉伸
- 登入時顯示置中 Modal（spinner + 進度條），完成後自動關閉

## 密碼設定頁
- 左右兩欄並排
- 左欄：網站登入密碼
- 右欄：對帳單檢視密碼

## 開發規則
- 金額欄位讀取後一律 parseFloat()
- 前端與後端分開修改，每次只改一個檔案
- Code.gs 修改後必須重新部署 Apps Script
- 每次完成後自動執行 git add . / git commit / git push origin main
- 每次任務最多 3～5 個項目，前後端分開處理