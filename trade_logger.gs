// Google Apps Script to maintain a trade input table and append edits to Ledger

const DATA_SHEET_NAME = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';
const TRADE_START_ROW = 20; // Row where trade headers begin
const TRADE_HEADERS = ['Symbol', 'Side', 'Quantity', 'Price', 'Trade Time', 'Note'];

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

/** Add headers when the spreadsheet is opened */
function onOpen(e) {
  ensureTradeArea();
  ensureLedgerSheet();
}

/** Monitor edits and route to the appropriate handler */
function onEdit(e) {
  try {
    if (!e) return;
    var sheet = e.range.getSheet();
    var name = sheet.getName();

    if (name === DATA_SHEET_NAME) {
      handleDataEdit(e);
    } else if (name === LEDGER_SHEET_NAME) {
      ledgerOnEdit(e); // defined in ledger.gs
    }
  } catch (err) {
    Logger.log(err);
  }
}

/** Append completed trade rows from the Data sheet to the Ledger */
function handleDataEdit(e) {
  var range = e.range;
  var row = range.getRow();
  if (row <= TRADE_START_ROW) return;

  var sheet = range.getSheet();
  var values = sheet.getRange(row, 1, 1, TRADE_HEADERS.length).getValues()[0];
  if (!values[0] || !values[1] || !values[2] || !values[3] || !values[4]) {
    return; // require all mandatory fields
  }

  var ledger = ensureLedgerSheet();
  ledger.appendRow(values.concat([new Date()]));
}

/** Find price for symbol at the given timestamp in Data sheet */
function findPrice(symbol, timeStr) {
  var sheet = getSheet(); // from coinbase_2h.gs
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  var timeVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var col = {BTC:2, ETH:3, SOL:4}[symbol];
  if (!col) return null;
  for (var i = 0; i < timeVals.length; i++) {
    if (timeVals[i][0] == timeStr) {
      return sheet.getRange(i + 2, col).getValue();
    }
  }
  return null;
}

/** Return the latest price for each symbol from the Data sheet */
function getLatestPrices() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  var prices = {BTC:0, ETH:0, SOL:0};
  if (lastRow <= 1) return prices;
  prices.BTC = sheet.getRange(lastRow, 2).getValue();
  prices.ETH = sheet.getRange(lastRow, 3).getValue();
  prices.SOL = sheet.getRange(lastRow, 4).getValue();
  return prices;
}

/** Recalculate ledger formulas for all rows and update summary */
function recomputeLedger() {
  var sheet = getLedgerSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var col = {id:0, time:1, sym:2, side:3, price:4, qty:5,
             amt:6, pos:7, avg:8, pnl:9};

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
