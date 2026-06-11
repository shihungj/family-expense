// ============================================================
// 家庭每月支出明細系統 - Google Apps Script 後端
// ============================================================

const SHEET_TRANSACTIONS    = 'Transactions';
const SHEET_BANK_SETTINGS   = 'BankSettings';
const SHEET_MERCHANT_RULES  = 'MerchantRules';
const SHEET_SYSTEM_SETTINGS = 'SystemSettings';

const EXPECTED_API_KEY = 'family-expense-2026';

// Fallback bank code map (superseded by col C of BankSettings)
const BANK_ABBREV = {
  '玉山銀行': 'ESUN',
  '中國信託': 'CTBC',
  '富邦銀行': 'FUBO',
  '永豐銀行': 'SINO',
  '聯邦銀行': 'UNIO',
};

// ── CORS ──────────────────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params   = e.parameter || {};
  const postData = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
  const action   = params.action || postData.action;

  const apiKey = params.apiKey || postData.apiKey;
  if (!verifyApiKey(apiKey)) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  let result;
  try {
    switch (action) {
      case 'login':               result = login(postData);              break;
      case 'changePassword':      result = changePassword(postData);     break;
      case 'getTransactions':     result = getTransactions(params);      break;
      case 'addTransactions':     result = addTransactions(postData);    break;
      case 'addTransaction':      result = addTransaction(postData);     break;  // [Bug2] 手動新增
      case 'deleteTransactions':  result = deleteTransactions(postData); break;
      case 'updateAttribution':   result = updateAttribution(postData);  break;
      case 'getKPI':              result = getKPI(params);               break;
      case 'getBankSettings':     result = getBankSettings();            break;
      case 'addBankSetting':      result = addBankSetting(postData);     break;  // [Bug3] 新增銀行
      case 'updateBankSetting':   result = updateBankSetting(postData);  break;  // [Bug3] 修改銀行
      case 'deleteBankSetting':   result = deleteBankSetting(postData);  break;  // [Bug3] 刪除銀行
      case 'getMerchantRules':    result = getMerchantRules();           break;
      case 'addMerchantRule':     result = addMerchantRule(postData);    break;  // [Bug6] 新增規則排第一
      case 'updateMerchantRule':  result = updateMerchantRule(postData); break;
      case 'deleteMerchantRule':  result = deleteMerchantRule(postData); break;
      case 'getBillingMonths':    result = getBillingMonths();           break;
      case 'getAvailableMonths':  result = getAvailableMonths();         break;
      case 'getStatementData':   result = getStatementData(postData);   break;
      case 'getSystemInfo':      result = getSystemInfo();              break;
      case 'backupData':         result = backupData();                 break;
      case 'restoreData':        result = restoreData(postData);        break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message, stack: err.stack };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyApiKey(key) { return key === EXPECTED_API_KEY; }

// ── Sheet helpers ─────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  if (name === SHEET_TRANSACTIONS) {
    sheet.appendRow(['入帳日', '消費日', '銀行', '消費明細', '金額', '歸屬', '帳單月份', '批次編號']);
    sheet.getRange('A:B').setNumberFormat('@');
    sheet.getRange('G:G').setNumberFormat('@');

  } else if (name === SHEET_BANK_SETTINGS) {
    // Columns: A=銀行, B=結帳日, C=銀行縮寫
    sheet.appendRow(['銀行', '結帳日', '銀行縮寫']);
    [
      ['玉山銀行', 13, 'ESUN'],
      ['永豐銀行', 26, 'SINO'],
      ['聯邦銀行', 12, 'UNIO'],
      ['中國信託', 25, 'CTBC'],
      ['富邦銀行', 26, 'FUBO'],
    ].forEach(r => sheet.appendRow(r));

  } else if (name === SHEET_MERCHANT_RULES) {
    sheet.appendRow(['關鍵字', '歸屬']);
    const defaults = [
      ['中油條碼_Autopass', '慧鳳應付'],
      ['0918169429',        '慧鳳應付'],
      ['悠遊卡',            '慧鳳應付'],
      ['連加*台灣中油條碼支付', '世鴻應付'],
      ['0939899529',        '世鴻應付'],
      ['ETC',               '世鴻應付'],
      ['遠通',              '世鴻應付'],
      ['家樂福',            '共同支付'],
      ['中華電信',          '共同支付'],
      ['台水',              '共同支付'],
      ['自來水',            '共同支付'],
      ['台電',              '共同支付'],
      ['電費',              '共同支付'],
      ['天然氣',            '共同支付'],
      ['瓦斯',              '共同支付'],
      ['其他',              '未分類'],
    ];
    defaults.forEach(r => sheet.appendRow(r));

  } else if (name === SHEET_SYSTEM_SETTINGS) {
    sheet.appendRow(['key', 'value']);
    sheet.appendRow(['account_shihung',  'shihung']);
    sheet.appendRow(['password_shihung', Utilities.base64Encode('0000')]);
    sheet.appendRow(['account_huifeng',  'huifeng']);
    sheet.appendRow(['password_huifeng', Utilities.base64Encode('0000')]);
  }
}

// ── Date helpers ──────────────────────────────────────────────
function toDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '/' +
      String(val.getMonth() + 1).padStart(2, '0') + '/' +
      String(val.getDate()).padStart(2, '0');
  }
  return String(val);
}

// ── 全形轉半形 ────────────────────────────────────────────────
function toHalfWidth(str) {
  return str.replace(/[！-～]/g, c =>
    String.fromCharCode(c.charCodeAt(0) - 0xFEE0)
  ).replace(/　/g, ' ');
}

// ── [Bug1] 穩健金額解析 ────────────────────────────────────────
// Google Sheets 有時以文字格式回傳數值，parseFloat('') 或 parseFloat(null) 回傳 NaN。
// parseAmount 統一轉成數字，若無法解析則回傳 0。
function parseAmount(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  const n = Number(raw);
  if (!isNaN(n)) return n;
  // 最後手段：移除非數字字元（如千分位逗號）後再解析
  const cleaned = parseFloat(String(raw).replace(/[^\d.-]/g, ''));
  return isNaN(cleaned) ? 0 : cleaned;
}

// ── Bank code map ─────────────────────────────────────────────
// Reads bankCode from BankSettings sheet (col C), falls back to hardcoded map.
function buildBankCodeMap() {
  const sheet = getSheet(SHEET_BANK_SETTINGS);
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    const bankName = String(data[i][0] || '').trim();
    const bankCode = String(data[i][2] || '').trim();
    if (bankName) {
      map[bankName] = bankCode || BANK_ABBREV[bankName] || '';
    }
  }
  return map;
}

function getBatchId(bankName, billingMonth, bankCodeMap) {
  const codeMap = bankCodeMap || {};
  const code    = codeMap[bankName] || BANK_ABBREV[bankName] || 'UNKN';
  const ym      = billingMonth.replace('/', '');
  return code + '-' + ym;
}

// ── Attribution via MerchantRules ─────────────────────────────
function classifyMerchant(detail) {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const keyword     = String(data[i][0]);
    const attribution = String(data[i][1]);
    if (keyword && detail.includes(keyword)) return attribution;
  }
  return '未分類';
}

// ── CRUD: Transactions ────────────────────────────────────────
// Column layout (0-indexed):
// 0:入帳日  1:消費日  2:銀行  3:消費明細  4:金額  5:歸屬  6:帳單月份  7:批次編號

function getTransactions(params) {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data  = sheet.getDataRange().getValues();
  const rows  = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    rows.push({
      rowIndex:     i + 1,
      postingDate:  toDateStr(row[0]),
      date:         toDateStr(row[1]),
      bank:         row[2],
      detail:       row[3],
      amount:       parseAmount(row[4]),   // [Bug1] 確保回傳純數字
      attribution:  row[5],
      billingMonth: String(row[6] || ''),
      batchId:      String(row[7] || ''),
    });
  }

  let filtered = rows;
  if (params.billingMonth) filtered = filtered.filter(r => r.billingMonth === params.billingMonth);
  if (params.bank)         filtered = filtered.filter(r => r.bank === params.bank);
  if (params.attribution)  filtered = filtered.filter(r => r.attribution === params.attribution);

  return { transactions: filtered };
}

function addTransactions(data) {
  const sheet        = getSheet(SHEET_TRANSACTIONS);
  const items        = data.transactions || [];
  const billingMonth = String(data.billingMonth || '');

  if (!billingMonth) return { success: false, error: '請選擇帳單月份' };
  if (!items.length)  return { success: true, added: 0 };

  const bankCodeMap = buildBankCodeMap();

  // Collect unique batch IDs
  const batchIdSet = new Set(
    items.map(item => getBatchId(String(item.bank), billingMonth, bankCodeMap))
  );

  // Check duplicates
  const existing = sheet.getDataRange().getValues();
  for (const batchId of batchIdSet) {
    for (let i = 1; i < existing.length; i++) {
      if (String(existing[i][7]) === batchId) {
        return { success: false, duplicate: true, batchId };
      }
    }
  }

  const rowsData = items.map(item => {
    const attribution = classifyMerchant(String(item.detail));
    const batchId     = getBatchId(String(item.bank), billingMonth, bankCodeMap);
    return [
      String(item.postingDate), String(item.date), String(item.bank),
      String(item.detail), parseAmount(item.amount),
      attribution, billingMonth, batchId,
    ];
  });

  const firstNewRow = sheet.getLastRow() + 1;
  const numRows     = rowsData.length;
  sheet.getRange(firstNewRow, 1, numRows, 1).setNumberFormat('@'); // 入帳日
  sheet.getRange(firstNewRow, 2, numRows, 1).setNumberFormat('@'); // 消費日
  sheet.getRange(firstNewRow, 7, numRows, 1).setNumberFormat('@'); // 帳單月份
  sheet.getRange(firstNewRow, 1, numRows, 8).setValues(rowsData);

  return { success: true, added: numRows };
}

// [Bug2] 手動單筆新增
function addTransaction(data) {
  const sheet        = getSheet(SHEET_TRANSACTIONS);
  const billingMonth = String(data.billingMonth || '');
  const bankName     = String(data.bank || '');

  if (!billingMonth) return { success: false, error: '請選擇帳單月份' };
  if (!data.postingDate || !data.date || !bankName || !data.detail || !data.amount) {
    return { success: false, error: '請填寫所有必填欄位' };
  }

  const bankCodeMap  = buildBankCodeMap();
  const batchId      = getBatchId(bankName, billingMonth, bankCodeMap);
  const attribution  = String(data.attribution || '未分類');

  const rowNum = sheet.getLastRow() + 1;
  sheet.getRange(rowNum, 1, 1, 1).setNumberFormat('@'); // 入帳日
  sheet.getRange(rowNum, 2, 1, 1).setNumberFormat('@'); // 消費日
  sheet.getRange(rowNum, 7, 1, 1).setNumberFormat('@'); // 帳單月份
  sheet.getRange(rowNum, 1, 1, 8).setValues([[
    String(data.postingDate), String(data.date), bankName,
    String(data.detail), parseAmount(data.amount),
    attribution, billingMonth, batchId,
  ]]);

  return { success: true };
}

// Batch or single delete
function deleteTransactions(data) {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const rowIndices = (data.rowIndices || [])
    .map(i => parseInt(i))
    .filter(i => !isNaN(i) && i >= 2)
    .sort((a, b) => b - a); // Descending to avoid index shifting

  for (const ri of rowIndices) {
    sheet.deleteRow(ri);
  }
  return { success: true, deleted: rowIndices.length };
}

function updateAttribution(data) {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  sheet.getRange(data.rowIndex, 6).setValue(data.attribution);
  return { success: true };
}

// ── KPI ────────────────────────────────────────────────────────
function getKPI(params) {
  const billingMonth = params.billingMonth;
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data  = sheet.getDataRange().getValues();

  // Compute previous billing month string (YYYY/MM)
  let lastMonth = '';
  if (billingMonth) {
    const parts = billingMonth.split('/');
    if (parts.length === 2) {
      let y = parseInt(parts[0]);
      let m = parseInt(parts[1]) - 1;
      if (m === 0) { m = 12; y -= 1; }
      lastMonth = y + '/' + String(m).padStart(2, '0');
    }
  }

  let shihong = 0, huifeng = 0, common = 0, total = 0;
  let totalLastMonth = 0;
  let transactionCount = 0, unclassifiedCount = 0, maxSingleAmount = 0;
  const importedBankSet = new Set();
  const bankAmountMap   = {};  // bankName -> total amount
  const bankCountMap    = {};  // bankName -> transaction count
  const paymentTools    = { linepay: 0, icashpay: 0, easycard: 0, easywallet: 0, carmoji: 0, creditcard: 0 };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;

    const rowMonth = String(row[6] || '');
    const amount   = parseAmount(row[4]);
    const attr     = String(row[5]);

    // Accumulate last month total
    if (lastMonth && rowMonth === lastMonth) {
      totalLastMonth += amount;
    }

    if (billingMonth && rowMonth !== billingMonth) continue;

    // Current month stats
    total += amount;
    transactionCount += 1;
    if (amount > maxSingleAmount) maxSingleAmount = amount;
    if (attr === '未分類') unclassifiedCount += 1;

    const bankName = String(row[2] || '').trim();
    if (bankName) {
      importedBankSet.add(bankName);
      bankAmountMap[bankName] = (bankAmountMap[bankName] || 0) + amount;
      bankCountMap[bankName]  = (bankCountMap[bankName]  || 0) + 1;
    }

    if      (attr === '世鴻應付') shihong += amount;
    else if (attr === '慧鳳應付') huifeng += amount;
    else if (attr === '共同支付') common  += amount;

    // Payment tool classification (case-insensitive)
    const detail = toHalfWidth(String(row[3] || '')).toLowerCase();
    if      (detail.includes('連加') || detail.includes('連支'))            paymentTools.linepay    += 1;
    else if (detail.includes('icash pay'))                                   paymentTools.icashpay   += 1;
    else if (detail.includes('悠遊卡'))                                      paymentTools.easycard   += 1;
    else if (detail.includes('悠遊付'))                                      paymentTools.easywallet += 1;
    else if (detail.includes('中油條碼_autopass'))                           paymentTools.carmoji    += 1;
    else                                                                     paymentTools.creditcard += 1;
  }

  // Build notImportedBanks from BankSettings
  const bankSheet = getSheet(SHEET_BANK_SETTINGS);
  const bankData  = bankSheet.getDataRange().getValues();
  const allBanks  = [];
  for (let i = 1; i < bankData.length; i++) {
    const name = String(bankData[i][0] || '').trim();
    if (name) allBanks.push(name);
  }
  const importedBanks    = Array.from(importedBankSet);
  const notImportedBanks = allBanks.filter(b => !importedBankSet.has(b));

  // topBankByAmount / topBankByCount
  let topBankByAmount = '';
  let topBankByCount  = '';
  let maxBankAmount   = -1;
  let maxBankCount    = -1;
  for (const bank in bankAmountMap) {
    if (bankAmountMap[bank] > maxBankAmount) { maxBankAmount = bankAmountMap[bank]; topBankByAmount = bank; }
    if (bankCountMap[bank]  > maxBankCount)  { maxBankCount  = bankCountMap[bank];  topBankByCount  = bank; }
  }

  return {
    shihong:           (Math.round(shihong) + Math.ceil(common / 2))  || 0,
    huifeng:           (Math.round(huifeng) + Math.floor(common / 2)) || 0,
    common:            Math.round(common)          || 0,
    total:             Math.round(total)            || 0,
    totalLastMonth:    Math.round(totalLastMonth)   || 0,
    transactionCount:  transactionCount,
    unclassifiedCount: unclassifiedCount,
    importedBanks:     importedBanks,
    notImportedBanks:  notImportedBanks,
    paymentTools:      paymentTools,
    maxSingleAmount:   Math.round(maxSingleAmount)  || 0,
    topBankByAmount:   topBankByAmount,
    topBankByCount:    topBankByCount,
  };
}

// ── [Bug3][Bug4] Bank Settings ─────────────────────────────────
// BankSettings columns: A=銀行, B=結帳日, C=銀行縮寫
// [Bug4] 自動補齊 col C（既有表格無此欄的一次性遷移邏輯）
function getBankSettings() {
  const sheet = getSheet(SHEET_BANK_SETTINGS);

  // [Bug4] 確保 C1 標題存在（舊表格可能沒有）
  if (!sheet.getRange(1, 3).getValue()) {
    sheet.getRange(1, 3).setValue('銀行縮寫');
  }

  const data  = sheet.getDataRange().getValues();
  const banks = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    const bankName = String(data[i][0]).trim();
    let   bankCode = String(data[i][2] || '').trim();

    // [Bug4] col C 為空時，從硬編碼對照表取值並自動回寫（一次性遷移）
    if (!bankCode && BANK_ABBREV[bankName]) {
      bankCode = BANK_ABBREV[bankName];
      sheet.getRange(i + 1, 3).setValue(bankCode);
    }

    banks.push({
      rowIndex:  i + 1,
      bank:      bankName,
      cutoffDay: parseInt(data[i][1]) || 1,
      bankCode:  bankCode,
    });
  }
  return { banks };
}

// [Bug3] 新增銀行
function addBankSetting(data) {
  const sheet = getSheet(SHEET_BANK_SETTINGS);
  // Check duplicate name
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (String(existing[i][0]) === String(data.bank)) {
      return { success: false, error: '銀行名稱已存在' };
    }
  }
  sheet.appendRow([String(data.bank), parseInt(data.cutoffDay) || 1, String(data.bankCode || '')]);
  return { success: true };
}

// [Bug3] 修改銀行
function updateBankSetting(data) {
  const sheet    = getSheet(SHEET_BANK_SETTINGS);
  const rowIndex = parseInt(data.rowIndex);
  if (data.cutoffDay !== undefined) {
    sheet.getRange(rowIndex, 2).setValue(parseInt(data.cutoffDay) || 1);
  }
  if (data.bankCode !== undefined) {
    sheet.getRange(rowIndex, 3).setValue(String(data.bankCode));
  }
  return { success: true };
}

// [Bug3] 刪除銀行
function deleteBankSetting(data) {
  const sheet = getSheet(SHEET_BANK_SETTINGS);
  sheet.deleteRow(parseInt(data.rowIndex));
  return { success: true };
}

// ── Available months (distinct billing months, descending) ────
function getAvailableMonths() {
  return getBillingMonths();
}

// ── Billing months list ───────────────────────────────────────
function getBillingMonths() {
  const sheet  = getSheet(SHEET_TRANSACTIONS);
  const data   = sheet.getDataRange().getValues();
  const months = new Set();
  for (let i = 1; i < data.length; i++) {
    const m = String(data[i][6] || '');
    if (m) months.add(m);
  }
  return { months: Array.from(months).sort().reverse() };
}

// ── Merchant Rules ────────────────────────────────────────────
function getMerchantRules() {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  const data  = sheet.getDataRange().getValues();
  const rules = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) rules.push({ rowIndex: i + 1, keyword: data[i][0], attribution: data[i][1] });
  }
  return { rules };
}

// [Bug6] 新增規則插在第一筆（標題列下方）
function addMerchantRule(data) {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  sheet.insertRowBefore(2);                                     // 在標題列後插入空列
  sheet.getRange(2, 1, 1, 2).setValues([[data.keyword, data.attribution]]);
  return { success: true };
}

function updateMerchantRule(data) {
  const sheet    = getSheet(SHEET_MERCHANT_RULES);
  const rowIndex = parseInt(data.rowIndex);
  sheet.getRange(rowIndex, 1, 1, 2).setValues([[String(data.keyword), String(data.attribution)]]);
  return { success: true };
}

function deleteMerchantRule(data) {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  sheet.deleteRow(parseInt(data.rowIndex));
  return { success: true };
}

// ── System Info ───────────────────────────────────────────────
function getSystemInfo() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet(SHEET_SYSTEM_SETTINGS);

  // frontendVersion: 直接讀取 D2 儲存格（C1=frontendVersion 標題，D2 為對應值）
  const frontendVersion = String(sheet.getRange(2, 4).getValue() || '').trim();

  // lastUpdated: 試用 getLastUpdated()，不支援時 fallback 現在時間
  let lastUpdatedDate;
  try {
    lastUpdatedDate = ss.getLastUpdated();
  } catch (e) {
    lastUpdatedDate = new Date();
  }
  const lu = lastUpdatedDate;
  const lastUpdated =
    lu.getFullYear() + '/' +
    String(lu.getMonth() + 1).padStart(2, '0') + '/' +
    String(lu.getDate()).padStart(2, '0') + ' ' +
    String(lu.getHours()).padStart(2, '0') + ':' +
    String(lu.getMinutes()).padStart(2, '0') + ':' +
    String(lu.getSeconds()).padStart(2, '0');

  const txSheet = getSheet(SHEET_TRANSACTIONS);
  const txData  = txSheet.getDataRange().getValues();
  const months  = new Set();
  for (let i = 1; i < txData.length; i++) {
    const m = String(txData[i][6] || '');
    if (m) months.add(m);
  }

  return {
    frontendVersion,
    lastUpdated,
    sheetName:        ss.getName(),
    transactionCount: Math.max(0, txData.length - 1),
    billingMonths:    months.size,
  };
}

// ── Backup / Restore ──────────────────────────────────────────
function backupData() {
  const sheetNames = [
    SHEET_TRANSACTIONS,
    SHEET_MERCHANT_RULES,
    SHEET_BANK_SETTINGS,
    SHEET_SYSTEM_SETTINGS,
  ];
  const backup = {};
  sheetNames.forEach(name => {
    const sheet = getSheet(name);
    backup[name] = sheet.getDataRange().getValues();
  });
  return { success: true, data: backup };
}

function restoreData(payload) {
  const backup = payload.data;
  if (!backup) return { success: false, error: '缺少 data 欄位' };

  const sheetNames = [
    SHEET_TRANSACTIONS,
    SHEET_MERCHANT_RULES,
    SHEET_BANK_SETTINGS,
    SHEET_SYSTEM_SETTINGS,
  ];

  sheetNames.forEach(name => {
    const rows = backup[name];
    if (!rows || !rows.length) return;

    const sheet = getSheet(name);

    if (name === SHEET_SYSTEM_SETTINGS) {
      // 保留 C、D 欄現有值，只覆蓋 A、B 欄
      const currentData  = sheet.getDataRange().getValues();
      const currentRows  = currentData.length;
      const restoreRows  = rows.length;
      const maxRows      = Math.max(currentRows, restoreRows);

      // 先確保工作表有足夠列數
      const needed = maxRows - sheet.getLastRow();
      if (needed > 0) {
        sheet.insertRowsAfter(sheet.getLastRow(), needed);
      }

      // 逐列寫入 A、B 欄
      for (let i = 0; i < restoreRows; i++) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, 1).setValue(rows[i][0] !== undefined ? rows[i][0] : '');
        sheet.getRange(rowNum, 2).setValue(rows[i][1] !== undefined ? rows[i][1] : '');
      }
      // 若原本列數多於還原列數，清除多餘 A、B 欄
      for (let i = restoreRows; i < currentRows; i++) {
        sheet.getRange(i + 1, 1).clearContent();
        sheet.getRange(i + 1, 2).clearContent();
      }
    } else {
      sheet.clearContents();
      sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    }
  });

  return { success: true };
}

// ── Statement Data ────────────────────────────────────────────
// params: { month: 'YYYY/MM', person: 'shihong' | 'huifeng' }
function getStatementData(params) {
  const month  = String(params.month  || '');
  const person = String(params.person || '');

  if (!month)  return { error: '請提供帳單月份 (month)' };
  if (person !== 'shihong' && person !== 'huifeng') {
    return { error: 'person 必須為 shihong 或 huifeng' };
  }

  const selfAttr = person === 'shihong' ? '世鴻應付' : '慧鳳應付';

  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data  = sheet.getDataRange().getValues();

  const selfItems   = [];
  const commonItems = [];
  let   selfRaw     = 0;
  let   commonRaw   = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    if (String(row[6] || '') !== month) continue;

    const amount = parseAmount(row[4]);
    const attr   = String(row[5] || '');
    const item   = {
      date:   toDateStr(row[1]),
      bank:   String(row[2] || ''),
      detail: String(row[3] || ''),
      amount: amount,
    };

    if (attr === selfAttr) {
      selfItems.push(item);
      selfRaw += amount;
    } else if (attr === '共同支付') {
      commonItems.push(item);
      commonRaw += amount;
    }
  }

  const commonShare = person === 'shihong'
    ? Math.ceil(commonRaw / 2)
    : Math.floor(commonRaw / 2);

  // topBank: self items 中金額最高的銀行
  const bankAmountMap = {};
  for (const item of selfItems) {
    if (item.bank) bankAmountMap[item.bank] = (bankAmountMap[item.bank] || 0) + item.amount;
  }
  let topBank = '';
  let topAmt  = -1;
  for (const bank in bankAmountMap) {
    if (bankAmountMap[bank] > topAmt) { topAmt = bankAmountMap[bank]; topBank = bank; }
  }

  // maxSingle: self items 中最高單筆
  let maxSingle = 0;
  for (const item of selfItems) {
    if (item.amount > maxSingle) maxSingle = item.amount;
  }

  return {
    selfItems,
    commonItems,
    selfTotal:        Math.round(selfRaw),
    commonTotal:      Math.round(commonRaw),
    commonShare:      commonShare,
    grandTotal:       Math.round(selfRaw) + commonShare,
    transactionCount: selfItems.length,
    maxSingle:        Math.round(maxSingle),
    topBank:          topBank,
  };
}

// ── Auth: Login ───────────────────────────────────────────────
function login(data) {
  const sheet = getSheet(SHEET_SYSTEM_SETTINGS);
  const rows  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) map[String(rows[i][0])] = String(rows[i][1]);
  }

  const username = String(data.username || '');
  const password = String(data.password || '');

  if (!map['account_' + username]) return { success: false };

  const encodedPwd = map['password_' + username] || '';
  const decoded    = Utilities.newBlob(Utilities.base64Decode(encodedPwd)).getDataAsString();
  return decoded === password ? { success: true, username } : { success: false };
}

// ── Auth: Change Password ─────────────────────────────────────
function changePassword(data) {
  const sheet   = getSheet(SHEET_SYSTEM_SETTINGS);
  const rows    = sheet.getDataRange().getValues();
  const username        = String(data.username        || '');
  const currentPassword = String(data.currentPassword || '');
  const newPassword     = String(data.newPassword     || '');

  let pwdRowNum    = -1;
  let storedEncoded = '';

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 'password_' + username) {
      pwdRowNum     = i + 1;
      storedEncoded = String(rows[i][1]);
    }
  }

  if (pwdRowNum === -1) return { success: false, error: '帳號不存在' };

  const decoded = Utilities.newBlob(Utilities.base64Decode(storedEncoded)).getDataAsString();
  if (decoded !== currentPassword) return { success: false, error: '目前密碼錯誤' };

  sheet.getRange(pwdRowNum, 2).setValue(Utilities.base64Encode(newPassword));
  return { success: true };
}
