// ============================================================
// 家庭每月支出明細系統 - Google Apps Script 後端
// ============================================================

const SHEET_TRANSACTIONS = 'Transactions';
const SHEET_BANK_SETTINGS = 'BankSettings';
const SHEET_MERCHANT_RULES = 'MerchantRules';

// ── CORS ──────────────────────────────────────────────────────
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const postData = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
  const action = params.action || postData.action;

  let result;
  try {
    switch (action) {
      case 'getTransactions':    result = getTransactions(params);      break;
      case 'addTransactions':    result = addTransactions(postData);    break;
      case 'updateAttribution':  result = updateAttribution(postData);  break;
      case 'getKPI':             result = getKPI(params);               break;
      case 'getBankSettings':    result = getBankSettings();            break;
      case 'getMerchantRules':   result = getMerchantRules();           break;
      case 'addMerchantRule':    result = addMerchantRule(postData);    break;
      case 'deleteMerchantRule': result = deleteMerchantRule(postData); break;
      case 'reorderMerchantRule':result = reorderMerchantRule(postData);break;
      case 'getBillingMonths':   result = getBillingMonths();           break;
      case 'getRecentTransactions': result = getRecentTransactions();   break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message, stack: err.stack };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

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
    sheet.appendRow(['消費日', '銀行', '消費明細', '金額', '歸屬', '帳單月份']);
  } else if (name === SHEET_BANK_SETTINGS) {
    sheet.appendRow(['銀行', '結帳日']);
    [['玉山銀行', 13], ['永豐銀行', 26], ['聯邦銀行', 12], ['中國信託', 25], ['富邦銀行', 26]].forEach(r => sheet.appendRow(r));
  } else if (name === SHEET_MERCHANT_RULES) {
    sheet.appendRow(['關鍵字', '歸屬']);
    const defaults = [
      ['中油條碼_Autopass', '慧鳳應付'],
      ['0918169429', '慧鳳應付'],
      ['悠遊卡', '慧鳳應付'],
      ['連加*台灣中油條碼支付', '世鴻應付'],
      ['0939899529', '世鴻應付'],
      ['ETC', '世鴻應付'],
      ['遠通', '世鴻應付'],
      ['家樂福', '共同支付'],
      ['中華電信', '共同支付'],
      ['台水', '共同支付'],
      ['自來水', '共同支付'],
      ['台電', '共同支付'],
      ['電費', '共同支付'],
      ['天然氣', '共同支付'],
      ['瓦斯', '共同支付'],
      ['其他', '未分類'],
    ];
    defaults.forEach(r => sheet.appendRow(r));
  }
}

// ── Billing month calculation ─────────────────────────────────
function calcBillingMonth(dateStr, bankName) {
  const settings = getBankSettingsMap();
  const cutoff = settings[bankName] || 1;
  const parts = dateStr.split('/');
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);
  const day = parseInt(parts[2]);

  // day < cutoff → previous month billing; day >= cutoff → current month billing
  let billYear = year;
  let billMonth = month;
  if (day < cutoff) {
    billMonth = month - 1;
    if (billMonth === 0) { billMonth = 12; billYear = year - 1; }
  }
  return `${billYear}/${String(billMonth).padStart(2, '0')}`;
}

function getBankSettingsMap() {
  const sheet = getSheet(SHEET_BANK_SETTINGS);
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) map[data[i][0]] = parseInt(data[i][1]);
  }
  return map;
}

// ── Attribution via MerchantRules ─────────────────────────────
function classifyMerchant(detail) {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const keyword = String(data[i][0]);
    const attribution = String(data[i][1]);
    if (keyword && detail.includes(keyword)) return attribution;
  }
  return '未分類';
}

// ── CRUD: Transactions ────────────────────────────────────────
function getTransactions(params) {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    rows.push({
      rowIndex: i + 1,
      date: row[0],
      bank: row[1],
      detail: row[2],
      amount: row[3],
      attribution: row[4],
      billingMonth: row[5],
    });
  }

  let filtered = rows;
  if (params.billingMonth) filtered = filtered.filter(r => r.billingMonth === params.billingMonth);
  if (params.bank) filtered = filtered.filter(r => r.bank === params.bank);
  if (params.attribution) filtered = filtered.filter(r => r.attribution === params.attribution);

  return { transactions: filtered };
}

function getRecentTransactions() {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    rows.push({
      rowIndex: i + 1,
      date: row[0],
      bank: row[1],
      detail: row[2],
      amount: row[3],
      attribution: row[4],
      billingMonth: row[5],
    });
  }
  rows.sort((a, b) => String(b.date).localeCompare(String(a.date)));
  return { transactions: rows.slice(0, 50) };
}

function addTransactions(data) {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const items = data.transactions || [];
  let added = 0;
  for (const item of items) {
    const attribution = classifyMerchant(String(item.detail));
    const billingMonth = calcBillingMonth(String(item.date), String(item.bank));
    sheet.appendRow([item.date, item.bank, item.detail, item.amount, attribution, billingMonth]);
    added++;
  }
  return { success: true, added };
}

function updateAttribution(data) {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  sheet.getRange(data.rowIndex, 5).setValue(data.attribution);
  return { success: true };
}

// ── KPI ───────────────────────────────────────────────────────
function getKPI(params) {
  const billingMonth = params.billingMonth;
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data = sheet.getDataRange().getValues();

  let shihong = 0, huifeng = 0, common = 0, unclassifiedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (billingMonth && row[5] !== billingMonth) continue;
    const amount = parseFloat(row[3]) || 0;
    const attr = String(row[4]);
    if (attr === '世鴻應付') shihong += amount;
    else if (attr === '慧鳳應付') huifeng += amount;
    else if (attr === '共同支付') common += amount;
    else if (attr === '未分類') unclassifiedCount++;
  }

  return {
    shihong: Math.round(shihong + common / 2),
    huifeng: Math.round(huifeng + common / 2),
    common: Math.round(common),
    unclassified: unclassifiedCount,
  };
}

// ── Bank Settings ─────────────────────────────────────────────
function getBankSettings() {
  const sheet = getSheet(SHEET_BANK_SETTINGS);
  const data = sheet.getDataRange().getValues();
  const banks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) banks.push({ bank: data[i][0], cutoffDay: parseInt(data[i][1]) });
  }
  return { banks };
}

// ── Billing months list ───────────────────────────────────────
function getBillingMonths() {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  const months = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5]) months.add(String(data[i][5]));
  }
  const sorted = Array.from(months).sort().reverse();
  return { months: sorted };
}

// ── Merchant Rules ────────────────────────────────────────────
function getMerchantRules() {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  const data = sheet.getDataRange().getValues();
  const rules = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) rules.push({ rowIndex: i + 1, keyword: data[i][0], attribution: data[i][1] });
  }
  return { rules };
}

function addMerchantRule(data) {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  // Insert after header at position data.priority (1-based after header)
  const insertAt = data.priority ? parseInt(data.priority) + 1 : sheet.getLastRow() + 1;
  sheet.insertRowBefore(insertAt);
  sheet.getRange(insertAt, 1, 1, 2).setValues([[data.keyword, data.attribution]]);
  return { success: true };
}

function deleteMerchantRule(data) {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  sheet.deleteRow(parseInt(data.rowIndex));
  return { success: true };
}

function reorderMerchantRule(data) {
  const sheet = getSheet(SHEET_MERCHANT_RULES);
  const fromRow = parseInt(data.fromRowIndex);
  const toRow = parseInt(data.toRowIndex);
  const rowData = sheet.getRange(fromRow, 1, 1, 2).getValues();
  sheet.deleteRow(fromRow);
  const adjustedTo = toRow > fromRow ? toRow : toRow;
  sheet.insertRowBefore(adjustedTo);
  sheet.getRange(adjustedTo, 1, 1, 2).setValues(rowData);
  return { success: true };
}
