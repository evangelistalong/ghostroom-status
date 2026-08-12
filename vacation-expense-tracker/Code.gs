/**
 * US Vacation Expense Tracker — Samit Reimbursement
 *
 * Bound to the Google Sheet. Provides:
 *  - setupSpreadsheet(): one-time build of all tabs, dropdowns and summaries
 *  - A web app (doGet + API functions) that up to 3 people can use to log
 *    expenses and mark them reimbursed by Samit.
 *
 * Spreadsheet: https://docs.google.com/spreadsheets/d/198cFEHC7fvbEZHlBBjXBSGdU9G5rI0t8YDkOqlKVaDI/edit
 */

var SHEET_EXPENSES = 'Expenses';
var SHEET_SUMMARY = 'Monthly Summary';
var SHEET_SETTINGS = 'Settings';

var CATEGORIES = [
  'Hotels',
  'Plane Tickets',
  'Disney Tickets',
  'Universal Tickets',
  'Restaurants / Food',
  'Las Vegas Show Tickets',
  'Other'
];

var DEFAULT_PEOPLE = ['Person 1', 'Person 2', 'Person 3']; // edit on the Settings tab
var STATUSES = ['Pending', 'Reimbursed'];

var HEADERS = [
  'ID', 'Date', 'Category', 'Description', 'Amount (USD)', 'Paid By',
  'Reimburse Month', 'Status', 'Reimbursed On', 'Notes', 'Receipt',
  'Entered By', 'Added At'
];

var RECEIPT_FOLDER = 'Vacation Receipts (Samit Reimbursement)';

// Receipt reading (optional). Add an Anthropic API key under
// Project Settings > Script Properties as ANTHROPIC_API_KEY to switch it on.
// Without a key the app still works — you just type the amount yourself.
var CLAUDE_MODEL = 'claude-opus-5';
var CLAUDE_URL = 'https://api.anthropic.com/v1/messages';

// ---------------------------------------------------------------------------
// Menu + one-time setup
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Expense Tracker')
    .addItem('Run one-time setup', 'setupSpreadsheet')
    .addToUi();
}

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ----- Settings tab -----
  var settings = getOrCreateSheet_(ss, SHEET_SETTINGS);
  settings.clear();
  settings.getRange('A1').setValue('People (who can pay / log expenses)');
  settings.getRange('A2:A4').setValues(DEFAULT_PEOPLE.map(function (p) { return [p]; }));
  settings.getRange('C1').setValue('Categories');
  settings.getRange('C2:C' + (CATEGORIES.length + 1))
    .setValues(CATEGORIES.map(function (c) { return [c]; }));
  settings.getRange('E1').setValue('Reimbursed by');
  settings.getRange('E2').setValue('Samit');
  settings.getRange('A1:E1').setFontWeight('bold');
  settings.setColumnWidths(1, 5, 200);
  ss.setNamedRange('People', settings.getRange('A2:A4'));
  ss.setNamedRange('Categories', settings.getRange('C2:C' + (CATEGORIES.length + 1)));

  // ----- Expenses tab -----
  var exp = getOrCreateSheet_(ss, SHEET_EXPENSES);
  exp.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  var head = exp.getRange(1, 1, 1, HEADERS.length);
  head.setFontWeight('bold').setBackground('#1a3c6e').setFontColor('#ffffff');
  exp.setFrozenRows(1);
  exp.setColumnWidth(4, 260);  // Description
  exp.setColumnWidth(10, 220); // Notes

  var maxRows = exp.getMaxRows();
  exp.getRange(2, 3, maxRows - 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('Categories'), true).build());
  exp.getRange(2, 6, maxRows - 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('People'), true).build());
  exp.getRange(2, 8, maxRows - 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  exp.getRange(2, 2, maxRows - 1).setNumberFormat('yyyy-mm-dd');
  exp.getRange(2, 9, maxRows - 1).setNumberFormat('yyyy-mm-dd');
  exp.getRange(2, 5, maxRows - 1).setNumberFormat('$#,##0.00');

  // Color rows by status: green when Reimbursed, amber when Pending.
  var dataRange = exp.getRange(2, 1, maxRows - 1, HEADERS.length);
  exp.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="Reimbursed"')
      .setBackground('#d9ead3').setRanges([dataRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="Pending"')
      .setBackground('#fff2cc').setRanges([dataRange]).build()
  ]);

  // ----- Monthly Summary tab -----
  var sum = getOrCreateSheet_(ss, SHEET_SUMMARY);
  sum.clear();
  sum.getRange('A1').setValue('OWED BY SAMIT — BY REIMBURSEMENT MONTH');
  sum.getRange('A2').setFormula(
    '=IFERROR(QUERY(' + SHEET_EXPENSES + '!A2:M,' +
    '"select G, sum(E), count(A) where A is not null group by G order by G ' +
    'label G \'Month\', sum(E) \'Total (USD)\', count(A) \'# Expenses\'",0),"No expenses yet")');
  sum.getRange('E1').setValue('PENDING (not yet paid by Samit)');
  sum.getRange('E2').setFormula(
    '=IFERROR(QUERY(' + SHEET_EXPENSES + '!A2:M,' +
    '"select G, sum(E) where H = \'Pending\' group by G order by G ' +
    'label G \'Month\', sum(E) \'Still Owed (USD)\'",0),"Nothing pending")');
  sum.getRange('A12').setValue('BY CATEGORY');
  sum.getRange('A13').setFormula(
    '=IFERROR(QUERY(' + SHEET_EXPENSES + '!A2:M,' +
    '"select C, sum(E), count(A) where A is not null group by C order by sum(E) desc ' +
    'label C \'Category\', sum(E) \'Total (USD)\', count(A) \'# Expenses\'",0),"No expenses yet")');
  sum.getRange('E12').setValue('BY PAYER');
  sum.getRange('E13').setFormula(
    '=IFERROR(QUERY(' + SHEET_EXPENSES + '!A2:M,' +
    '"select F, sum(E) where A is not null group by F ' +
    'label F \'Paid By\', sum(E) \'Total (USD)\'",0),"No expenses yet")');
  sum.getRange('A24').setValue('GRAND TOTALS');
  sum.getRange('A25:B27').setValues([
    ['Total spent', ''], ['Reimbursed by Samit', ''], ['Still owed', '']
  ]);
  sum.getRange('B25').setFormula('=SUM(' + SHEET_EXPENSES + '!E2:E)');
  sum.getRange('B26').setFormula(
    '=SUMIF(' + SHEET_EXPENSES + '!H2:H,"Reimbursed",' + SHEET_EXPENSES + '!E2:E)');
  sum.getRange('B27').setFormula('=B25-B26');
  sum.getRange('B25:B27').setNumberFormat('$#,##0.00');
  sum.getRangeList(['A1', 'E1', 'A12', 'E12', 'A24']).setFontWeight('bold');
  sum.setColumnWidths(1, 7, 160);

  ss.setActiveSheet(exp);
  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\nNext: edit the 3 names on the Settings tab, then deploy the web app ' +
    '(Deploy > New deployment > Web app).');
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ---------------------------------------------------------------------------
// Web app
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('App')
    .setTitle('US Vacation Expenses — Samit Reimbursement')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Everything the UI needs in one round trip. */
function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var exp = ss.getSheetByName(SHEET_EXPENSES);
  var rows = [];
  if (exp && exp.getLastRow() > 1) {
    var values = exp.getRange(2, 1, exp.getLastRow() - 1, HEADERS.length).getValues();
    var tz = ss.getSpreadsheetTimeZone();
    rows = values
      .filter(function (r) { return r[0] !== ''; })
      .map(function (r) {
        return {
          id: r[0],
          date: fmtDate_(r[1], tz),
          category: r[2],
          description: r[3],
          amount: Number(r[4]) || 0,
          paidBy: r[5],
          reimburseMonth: r[6],
          status: r[7],
          reimbursedOn: fmtDate_(r[8], tz),
          notes: r[9],
          receiptUrl: String(r[10] || ''),
          enteredBy: r[11]
        };
      });
  }
  return {
    people: readColumn_(ss, 'People'),
    categories: readColumn_(ss, 'Categories'),
    scanner: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'),
    expenses: rows.reverse() // newest first
  };
}

/**
 * Reads a receipt photo with Claude's vision API and returns the amount, date,
 * merchant and a suggested category so the form can be pre-filled.
 *
 * Returns {ok: false, reason: ...} rather than throwing when reading fails —
 * a bad scan should never block someone from typing the expense in by hand.
 */
function scanReceipt(b64) {
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, reason: 'no_key' };
  if (!b64) return { ok: false, reason: 'no_image' };

  // Use the live category list so edits on the Settings tab are picked up here.
  var categories;
  try {
    categories = readColumn_(SpreadsheetApp.getActiveSpreadsheet(), 'Categories');
  } catch (err) {
    categories = CATEGORIES;
  }
  if (!categories.length) categories = CATEGORIES;

  var schema = {
    type: 'object',
    properties: {
      total: {
        type: 'number',
        description: 'The final total actually paid, including tax and tip. 0 if unreadable.'
      },
      currency: { type: 'string', description: 'ISO code, e.g. USD. Empty if unclear.' },
      date: { type: 'string', description: 'Date on the receipt as yyyy-mm-dd. Empty if unclear.' },
      merchant: { type: 'string', description: 'Business name. Empty if unclear.' },
      category: { type: 'string', enum: categories },
      confidence: { type: 'string', enum: ['high', 'medium', 'low'] },
      notes: {
        type: 'string',
        description: 'Short note only if something needs a human eye (split bill, ' +
          'unreadable total, foreign currency). Otherwise empty.'
      }
    },
    required: ['total', 'currency', 'date', 'merchant', 'category', 'confidence', 'notes'],
    additionalProperties: false
  };

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: 8192,
    // Structured outputs guarantee the reply parses; low effort keeps this
    // fast and cheap for what is a straightforward extraction task.
    output_config: { effort: 'low', format: { type: 'json_schema', schema: schema } },
    fallbacks: 'default',
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: b64 } },
        {
          type: 'text',
          text: 'This is a receipt or booking confirmation from a US vacation. ' +
            'Read the grand total actually paid — the final amount including tax and ' +
            'tip, not the subtotal or a per-person share. Report the date printed on ' +
            'the receipt, the merchant, and the best-fitting category. ' +
            'Leave a field empty rather than guessing at it, and set confidence to ' +
            '"low" if the image is blurry, cut off, or ambiguous.'
        }
      ]
    }]
  };

  var res = UrlFetchApp.fetch(CLAUDE_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': key,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'server-side-fallback-2026-07-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    return { ok: false, reason: 'api_error', detail: res.getContentText().slice(0, 300) };
  }

  var body = JSON.parse(res.getContentText());
  if (body.stop_reason === 'refusal') return { ok: false, reason: 'refused' };

  // Thinking blocks come first, so pick out the text block rather than content[0].
  var text = '';
  for (var i = 0; i < (body.content || []).length; i++) {
    if (body.content[i].type === 'text') { text = body.content[i].text; break; }
  }
  if (!text) return { ok: false, reason: 'empty' };

  var data;
  try {
    data = JSON.parse(text);
  } catch (err) {
    return { ok: false, reason: 'unparsed' };
  }
  data.ok = true;
  return data;
}

function addExpense(e) {
  if (!e || !e.date || !e.category || !(Number(e.amount) > 0) || !e.paidBy) {
    throw new Error('Date, category, a positive amount and "paid by" are required.');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var exp = ss.getSheetByName(SHEET_EXPENSES);
    var id = 'EXP-' + Utilities.formatDate(new Date(), 'UTC', 'yyMMddHHmmssSSS');
    var receiptUrl = '';
    if (e.receiptB64) {
      receiptUrl = saveReceipt_(e.receiptB64, id, e.category);
    }
    exp.appendRow([
      id,
      e.date,
      e.category,
      String(e.description || ''),
      Number(e.amount),
      e.paidBy,
      e.reimburseMonth || defaultReimburseMonth_(e.date),
      'Pending',
      '',
      String(e.notes || ''),
      receiptUrl,
      String(e.enteredBy || Session.getActiveUser().getEmail() || ''),
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm')
    ]);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

/** Toggle Pending <-> Reimbursed. Stamps/clears "Reimbursed On". */
function setStatus(id, status) {
  if (STATUSES.indexOf(status) === -1) throw new Error('Bad status: ' + status);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var exp = ss.getSheetByName(SHEET_EXPENSES);
    var row = findRowById_(exp, id);
    exp.getRange(row, 8).setValue(status);
    exp.getRange(row, 9).setValue(
      status === 'Reimbursed'
        ? Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd')
        : '');
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function deleteExpense(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var exp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EXPENSES);
    exp.deleteRow(findRowById_(exp, id));
    return getData();
  } finally {
    lock.releaseLock();
  }
}

/** Mark every pending expense in one reimbursement month as paid. */
function reimburseMonth(month) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var exp = ss.getSheetByName(SHEET_EXPENSES);
    var today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    var last = exp.getLastRow();
    if (last > 1) {
      var range = exp.getRange(2, 1, last - 1, HEADERS.length);
      var values = range.getValues();
      for (var i = 0; i < values.length; i++) {
        if (values[i][6] === month && values[i][7] === 'Pending') {
          values[i][7] = 'Reimbursed';
          values[i][8] = today;
        }
      }
      range.setValues(values);
    }
    return getData();
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Stores a receipt photo (base64 JPEG from the web app) in a Drive folder
 * and returns a view link. Files are link-viewable so all 3 users can open
 * them from the app. The folder ID is cached in Script Properties.
 */
function saveReceipt_(b64, expenseId, category) {
  var props = PropertiesService.getScriptProperties();
  var folder;
  var folderId = props.getProperty('receiptFolderId');
  if (folderId) {
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (err) {
      folder = null; // folder was deleted; recreate below
    }
  }
  if (!folder) {
    var existing = DriveApp.getFoldersByName(RECEIPT_FOLDER);
    folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(RECEIPT_FOLDER);
    props.setProperty('receiptFolderId', folder.getId());
  }
  var blob = Utilities.newBlob(
    Utilities.base64Decode(b64), 'image/jpeg',
    expenseId + ' - ' + category + '.jpg');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function findRowById_(sheet, id) {
  var ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1)).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  throw new Error('Expense not found: ' + id);
}

function readColumn_(ss, namedRange) {
  return ss.getRangeByName(namedRange).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(String);
}

/** Samit pays the month AFTER the expense: 2026-08-20 -> "2026-09". */
function defaultReimburseMonth_(dateStr) {
  var parts = String(dateStr).split('-');
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1 + 1, 1);
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
}

function fmtDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v || '');
}
