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
  "totalLastMonth": 0,
  "transactionCount": 0,
  "importedBanks": [],
  "notImportedBanks": [],
  "paymentTools": {
    "linepay": 0,
    "icashpay": 0,
    "easycard": 0,
    "easywallet": 0,
    "carmoji": 0,
    "creditcard": 0
  },
  "maxSingleAmount": 0,
  "topBankByAmount": "",
  "topBankByCount": ""
}
```

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

## 首頁版面結構
1. 左側固定選單
2. 頂部帳單月份選擇器（左右箭頭＋點擊數字跳出年月選擇器）
3. KPI 三張卡片（世鴻藍、慧鳳紅、共同綠）
4. 三欄等寬區塊：
   - 左：帳單焦點總覽（總支出、較上期增減、交易筆數）
   - 中：支付工具統計（六項）
   - 右：消費行為統計（最高單筆、最高額銀行、最多次數銀行）
5. 帳單匯入進度（五張卡片平均分配寬度，三色狀態）
6. 底部三張快捷卡片（帳單匯入、分類規則管理、帳單預覽）

## 配色規範
- 網頁主底色：#F7F9FC
- 左右留白：#E8EEF7
- 側邊欄底色：#F5F6F8
- 卡片背景：#FFFFFF
- 卡片邊框：#ECEFF5
- 世鴻：#1E75EC / #F0F6FF
- 慧鳳：#E53935 / #FFECEB
- 共同：#1B9A40 / #E8F7EC
- 主文字：#1A202C
- 次要文字：#718096
- 已匯入綠：#38A169
- 尚未匯入橘：#DD6B20
- 盡速匯入紅：#EB4B4B
- 選單高亮背景：#FFF5F2
- 選單高亮文字：#DD6B20

## 全部交易明細頁規格（待實作）
- 表格固定高度 600px，內部垂直捲動軸
- 標題列固定不隨捲動
- 欄位：項次、入帳日、消費日、銀行、消費明細、金額、歸屬、帳單月份、批次編號、操作
- 篩選與排序規則：
  - 銀行：下拉篩選＋排序
  - 消費明細：關鍵字篩選
  - 金額：排序
  - 歸屬：下拉篩選
  - 帳單月份：下拉篩選＋排序
  - 入帳日、消費日、操作欄：無篩選無排序
- 多選功能＋批次修改歸屬按鈕＋批次刪除按鈕

## 對帳單預覽頁規格（待實作）
- Tab 切換：慧鳳對帳單（預設）、世鴻對帳單
- 慧鳳：應付總金額＋慧鳳應付明細＋共同支付明細（含÷2捨去後金額）
- 世鴻：應付總金額＋世鴻應付明細＋共同支付明細（含÷2進位後金額）
- 匯出 PDF 使用 html2pdf.js
- PDF 檔名：慧鳳對帳單_202605.pdf / 世鴻對帳單_202605.pdf

## 開發規則
- 金額欄位讀取後一律 parseFloat()
- 前端與後端分開修改，每次只改一個檔案
- Code.gs 修改後必須重新部署 Apps Script
- 每次完成後自動執行 git add . / git commit / git push origin main
- 每次任務最多 3～5 個項目，前後端分開處理