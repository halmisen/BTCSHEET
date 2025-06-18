// Google Apps Script utilities for a paper trading ledger

/** Headers used in the Ledger sheet */
const LEDGER_HEADERS = [
  'Trade ID',
  'Trade Time',
  'Symbol',
  'Side',
  'Price',
  'Quantity',
  'Trade Amount',
  'Running Position',
  'Average Cost',
  'Floating P&L'
];

/** Return the Ledger sheet, creating it if necessary */
function getLedgerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Ledger');
  if (!sheet) {
    sheet = ss.insertSheet('Ledger');
    sheet.appendRow(LEDGER_HEADERS);
    setupLedgerValidation();
  }
  return sheet;
}

/** Setup dropdowns for Symbol and Side columns */
function setupLedgerValidation() {
  var sheet = getLedgerSheet();
  var last = sheet.getMaxRows() - 1;
  var symbolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['BTC','ETH','SOL'], true)
    .build();
  sheet.getRange(2, 3, last).setDataValidation(symbolRule);

  var sideRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Buy','Sell'], true)
    .build();
  sheet.getRange(2, 4, last).setDataValidation(sideRule);
}

/** Find price for symbol at the given timestamp in Data sheet */
function findPrice(symbol, timeStr) {
  var sheet = getSheet(); // Data sheet from coinbase_2h.gs
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
    var signedQty = qty * sign;

    var prevPos = pos[sym];
    var prevAvg = avg[sym];
    var newPos = prevPos + signedQty;
    var newAvg = prevAvg;
    if (sign > 0) { // buy
      newAvg = (prevAvg * Math.abs(prevPos) + price * qty) / Math.abs(newPos);
    } else { // sell
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

  // Write summary block
  var start = sheet.getLastRow() + 2;
  var out = [['Symbol','Position','Avg Cost','Floating P&L']];
  ['BTC','ETH','SOL'].forEach(function(s) {
    out.push([s, pos[s], avg[s], (lastPrices[s] - avg[s]) * pos[s]]);
  });
  sheet.getRange(start, 1, out.length, out[0].length).clearContent();
  sheet.getRange(start, 1, out.length, out[0].length).setValues(out);
}

/** Handle edits on the Ledger sheet */
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() != 'Ledger') return;
  var row = e.range.getRow();
  if (row <= 1) return;

  var values = sheet.getRange(row, 1, 1, LEDGER_HEADERS.length).getValues()[0];
  var id = values[0];
  var time = values[1];
  var sym = values[2];
  var side = values[3];
  var price = values[4];
  var qty = values[5];

  if (!id && (time && sym)) {
    var nextId = sheet.getRange(sheet.getLastRow(), 1).getValue();
    nextId = nextId ? nextId + 1 : 1;
    sheet.getRange(row, 1).setValue(nextId);
  }

  if (!price && time && sym) {
    var p = findPrice(sym, time);
    if (p != null) sheet.getRange(row, 5).setValue(p);
  }

  if (time && sym && side && qty) {
    recomputeLedger();
  }
}

/** Convenience to create the Ledger sheet and validation rules */
function initLedger() {
  var sheet = getLedgerSheet();
  sheet.clear();
  sheet.appendRow(LEDGER_HEADERS);
  setupLedgerValidation();
}
