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

/**
 * Handle edits on the Ledger sheet. Renamed from onEdit to avoid
 * conflicts with the global trigger in trade_logger.gs.
 */
function ledgerOnEdit(e) {
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
