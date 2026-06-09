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
      case 'addTransaction':      result = addTransaction(postData);     break;
      case 'deleteTransactions':  result = deleteTransactions(postData); break;
      case 'updateAttribution':   result = updateAttribution(postData);  break;
      case 'getKPI':              result = getKPI(params);               break;
      case 'getBankSettings':     result = getBankSettings();            break;
      case 'addBankSetting':      result = addBankSetting(postData);     break;
      case 'updateBankSetting':   result = updateBankSetting(postData);  break;
      case 'deleteBankSetting':   result = deleteBankSetting(postData);  break;
      case 'getMerchantRules':    result = getMerchantRules();           break;
      case 'addMerchantRule':     result = addMerchantRule(postData);    break;
      case 'updateMerchantRule':  result = updateMerchantRule(postData); break;
      case 'deleteMerchantRule':  result = deleteMerchantRule(postData); break;
      case 'getBillingMonths':    result = getBillingMonths();           break;
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

// ── Bank code map ─────────────────────────────────────────────
// Reads bankCode from BankSettings sheet (col C), falls back to hardcoded map.
function buildBankCodeMap() {
  const sheet = getSheet(SHEET_BANK_SETTINGS);
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][2]) {
      map[String(data[i][0])] = String(data[i][2]);
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
      amount:       row[4],
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
      String(item.detail), parseFloat(item.amount) || 0,
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

// Manual single-transaction entry
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
    String(data.detail), parseFloat(data.amount) || 0,
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

// ── KPI ───────────────────────────────────────────────────────
function getKPI(params) {
  const billingMonth = params.billingMonth;
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data  = sheet.getDataRange().getValues();

  let shihong = 0, huifeng = 0, common = 0, total = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    if (billingMonth && String(row[6] || '') !== billingMonth) continue;

    const amount = parseFloat(row[4]) || 0;
    const attr   = String(row[5]);
    total += amount;  // All transactions contribute to total

    if      (attr === '世鴻應付') shihong += amount;
    else if (attr === '慧鳳應付') huifeng += amount;
    else if (attr === '共同支付') common  += amount;
  }

  return {
    // 世鴻 = 世鴻原始 + ceil(共同/2)；慧鳳 = 慧鳳原始 + floor(共同/2)
    shihong: Math.round(shihong) + Math.ceil(common / 2),
    huifeng: Math.round(huifeng) + Math.floor(common / 2),
    common:  Math.round(common),
    total:   Math.round(total),
  };
}

// ── Bank Settings ─────────────────────────────────────────────
// BankSettings columns: A=銀行, B=結帳日, C=銀行縮寫
function getBankSettings() {
  const sheet = getSheet(SHEET_BANK_SETTINGS);
  const data  = sheet.getDataRange().getValues();
  const banks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      banks.push({
        rowIndex:  i + 1,
        bank:      String(data[i][0]),
        cutoffDay: parseInt(data[i][1]) || 1,
        bankCode:  String(data[i][2] || ''),
      });
    }
  }
  return { banks };
}

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

function deleteBankSetting(data) {
  const sheet = getSheet(SHEET_BANK_SETTINGS);
  sheet.deleteRow(parseInt(data.rowIndex));
  return { success: true };
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

// Always inserts at the first position (row 2), pushing existing rules down.
function addMerchantRule(data) {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  sheet.insertRowBefore(2); // Insert after header row
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
