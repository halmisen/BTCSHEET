// Google Apps Script to maintain a trade input table and append edits to Ledger

const DATA_SHEET_NAME = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';
const TRADE_START_ROW = 20; // Row where trade headers begin
const TRADE_HEADERS = ['Trade ID', 'Trade Time', 'Symbol', 'Side', 'Price', 'Quantity', 'Note'];

// Ledger columns mirroring ledger.gs but with Note and Logged Timestamp fields
const LEDGER_HEADERS = [
  'Trade ID',
  'Trade Time',
  'Symbol',
  'Side',
  'Price',
  'Quantity',
  'Note',
  'Trade Amount',
  'Running Position',
  'Average Cost',
  'Floating P&L',
  'Logged Timestamp'
];

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
    try {
      // Copy formatting from the sheet's first header row or row 20 if it exists
      sheet.getRange(1, 1, 1, TRADE_HEADERS.length)
        .copyFormatToRange(sheet, 1, TRADE_HEADERS.length,
                           TRADE_START_ROW, TRADE_START_ROW);
    } catch (err) {}
    var templates = 5; // clear at least five blank rows
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
  var expected = LEDGER_HEADERS;
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
    var nextId = ledger.getRange(ledger.getLastRow(), 1).getValue();
    nextId = nextId ? nextId + 1 : 1;
    var rowData = [
      nextId,            // Trade ID
      values[4],         // Trade Time
      values[0],         // Symbol
      values[1],         // Side
      values[3],         // Price
      values[2],         // Quantity
      values[5],         // Note
      '', '', '', '',    // Calculated fields
      new Date()         // Logged Timestamp
    ];
    ledger.appendRow(rowData);
    recomputeLedger();
  } catch (err) {
    Logger.log(err);
  }
}

/** Recalculate ledger amounts and running stats for all rows */
function recomputeLedger() {
  var sheet = ensureLedgerSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var col = {id:0, time:1, sym:2, side:3, price:4, qty:5,
             note:6, amt:7, pos:8, avg:9, pnl:10};

  var pos = {BTC:0, ETH:0, SOL:0};
  var avg = {BTC:0, ETH:0, SOL:0};
  var lastPrices = getLatestPrices();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var sym = row[col.sym];
    if (!sym) continue;
    var price = parseFloat(row[col.price]);
    var qty = parseFloat(row[col.qty]);
    if (isNaN(price) || isNaN(qty)) continue;
    var sign = row[col.side] == 'Buy' ? 1 : -1;

    var prevPos = pos[sym];
    var prevAvg = avg[sym];
    var newPos = prevPos + qty * sign;
    var newAvg = prevAvg;
    if (sign > 0) {
      newAvg = (prevAvg * Math.abs(prevPos) + price * qty) / Math.abs(newPos);
    } else {
      if (Math.sign(prevPos) == Math.sign(newPos) && prevPos != 0) {
        newAvg = prevAvg;
      } else if (newPos == 0) {
        newAvg = 0;
      } else {
        newAvg = price;
      }
    }

    pos[sym] = newPos;
    avg[sym] = newAvg;

    row[col.amt] = price * qty * sign;
    row[col.pos] = newPos;
    row[col.avg] = newAvg;
    row[col.pnl] = (lastPrices[sym] - newAvg) * newPos;

    sheet.getRange(i + 1, col.amt + 1, 1, 4)
         .setValues([[row[col.amt], row[col.pos], row[col.avg], row[col.pnl]]]);
  }

  var start = sheet.getLastRow() + 2;
  var out = [['Symbol','Position','Avg Cost','Floating P&L']];
  ['BTC','ETH','SOL'].forEach(function(s) {
    out.push([s, pos[s], avg[s], (lastPrices[s] - avg[s]) * pos[s]]);
  });
  sheet.getRange(start, 1, out.length, out[0].length).clearContent();
  sheet.getRange(start, 1, out.length, out[0].length).setValues(out);
}
