// Google Apps Script to maintain a trade input table and append edits to Ledger

const DATA_SHEET_NAME = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';
const TRADE_START_ROW = 20; // Row where trade headers begin
const TRADE_HEADERS = ['Trade ID', 'Trade Time', 'Symbol', 'Side', 'Price', 'Quantity', 'Note'];

/** Ensure trade area headers exist and add blank template rows */
function ensureTradeArea() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return;
  var headerRange = sheet.getRange(TRADE_START_ROW, 1, 1, TRADE_HEADERS.length);
  var existing = headerRange.getValues()[0];
  var match = true;
  for (var i = 0; i < TRADE_HEADERS.length; i++) {
    if (existing[i].toString().trim() !== TRADE_HEADERS[i]) {
      match = false;
      break;
    }
  }
  if (!match) {
    headerRange.setValues([TRADE_HEADERS]);
    var templates = 5;
    sheet.getRange(TRADE_START_ROW + 1, 1, templates, TRADE_HEADERS.length)
         .clearContent();
  }
}

/** Ensure Ledger sheet exists with headers */
function ensureLedgerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LEDGER_SHEET_NAME);
  }
  var expected = TRADE_HEADERS.concat(['Logged Timestamp']);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expected);
  } else {
    var cur = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
    var ok = true;
    for (var j = 0; j < expected.length; j++) {
      if (cur[j].toString().trim() !== expected[j]) {
        ok = false;
        break;
      }
    }
    if (!ok) {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    }
  }
  return sheet;
}

/** Get the next available Trade ID from the Ledger sheet */
function getNextTradeId() {
  var sheet = ensureLedgerSheet();
  var lastRow = sheet.getLastRow();
  for (var r = lastRow; r >= 2; r--) {
    var id = sheet.getRange(r, 1).getValue();
    if (id !== '' && id != null) {
      return (parseInt(id, 10) || 0) + 1;
    }
  }
  return 1;
}

/** Add headers when the spreadsheet is opened */
function onOpen(e) {
  ensureTradeArea();
  ensureLedgerSheet();
}

/** Monitor edits in the trade area and append to Ledger */
function onEdit(e) {
  try {
    if (!e) return;
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() !== DATA_SHEET_NAME) return;

    var row = range.getRow();
    if (row <= TRADE_START_ROW) return;

    var rowRange = sheet.getRange(row, 1, 1, TRADE_HEADERS.length);
    var values = rowRange.getValues()[0];

    var symbol = values[2];
    var side = values[3];
    var price = values[4];
    var qty = values[5];

    if (!symbol || !side || !price || !qty) {
      return; // require trade details
    }

    if (!values[0]) {
      values[0] = getNextTradeId();
      sheet.getRange(row, 1).setValue(values[0]);
    }

    if (!values[1]) {
      values[1] = new Date();
      sheet.getRange(row, 2).setValue(values[1]);
    }

    var ledger = ensureLedgerSheet();
    ledger.appendRow(values.concat([new Date()]));
  } catch (err) {
    Logger.log(err);
  }
}
