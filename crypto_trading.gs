// Unified Google Apps Script for simulated crypto trading

// Sheet and table configuration
const DATA_SHEET_NAME   = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';
const TRADE_HEADER_ROW  = 20;                               // trade table header row
const TRADE_HEADERS     = ['Symbol','Side','Quantity','Price','Trade Time','Note'];
const LEDGER_HEADERS    = ['Trade ID','Trade Time','Symbol','Side','Price','Quantity',
                           'Trade Amount','Running Position','Average Cost','Floating P&L'];

/** Ensure trade input headers exist in the Data sheet */
function ensureTradeArea() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return;
  const rng = sheet.getRange(TRADE_HEADER_ROW, 1, 1, TRADE_HEADERS.length);
  const values = rng.getValues()[0];
  let match = true;
  for (let i = 0; i < TRADE_HEADERS.length; i++) {
    if (values[i] !== TRADE_HEADERS[i]) { match = false; break; }
  }
  if (!match) {
    rng.setValues([TRADE_HEADERS]);
    sheet.getRange(TRADE_HEADER_ROW + 1, 1, 5, TRADE_HEADERS.length).clearContent();
  }
}

/** Ensure the Ledger sheet exists and has the correct headers */
function ensureLedgerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(LEDGER_SHEET_NAME);
  const rng = sheet.getRange(1, 1, 1, LEDGER_HEADERS.length);
  const cur = rng.getValues()[0];
  let match = true;
  for (let i = 0; i < LEDGER_HEADERS.length; i++) {
    if (cur[i] !== LEDGER_HEADERS[i]) { match = false; break; }
  }
  if (!match) {
    sheet.clear();
    rng.setValues([LEDGER_HEADERS]);
  }
  return sheet;
}

/** Get latest price for every coin column in the Data sheet */
function getLatestPrices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1).getValues()[0];
  const prices = {};
  if (lastRow > 1) {
    headers.forEach((h, i) => {
      if (h) prices[h] = sheet.getRange(lastRow, i + 2).getValue();
    });
  }
  return prices;
}

/** Build the entire ledger based on trades in the Data sheet */
function syncLedgerWithData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  const ledgerSheet = ensureLedgerSheet();

  const lastRow = dataSheet.getLastRow();
  const rows = lastRow > TRADE_HEADER_ROW
    ? dataSheet.getRange(TRADE_HEADER_ROW + 1, 1, lastRow - TRADE_HEADER_ROW, TRADE_HEADERS.length).getValues()
    : [];

  const latest = getLatestPrices();
  const pos = {};          // running position per symbol
  const avg = {};          // average cost per symbol
  const ledgerRows = [];

  rows.forEach(r => {
    const [symbol, side, qtyVal, priceVal, timeVal, note] = r;
    const qty = parseFloat(qtyVal);
    const price = parseFloat(priceVal);
    if (!symbol || !side || isNaN(qty) || isNaN(price)) return; // skip incomplete rows
    const time = timeVal || new Date();

    if (pos[symbol] === undefined) { pos[symbol] = 0; avg[symbol] = 0; }
    const sign = side.toString().toLowerCase() === 'buy' ? 1 : -1;

    const prevPos = pos[symbol];
    const prevAvg = avg[symbol];
    const newPos  = prevPos + qty * sign;
    let   newAvg  = prevAvg;
    if (sign > 0) {
      newAvg = (prevAvg * Math.abs(prevPos) + price * qty) / Math.abs(newPos);
    } else {
      if (Math.sign(prevPos) === Math.sign(newPos) && prevPos !== 0) {
        newAvg = prevAvg;
      } else if (newPos === 0) {
        newAvg = 0;
      } else {
        newAvg = price;
      }
    }
    pos[symbol] = newPos;
    avg[symbol] = newAvg;

    const tradeAmt = price * qty * sign;
    const floatPnl = ((latest[symbol] || 0) - newAvg) * newPos;
    ledgerRows.push([ledgerRows.length + 1, time, symbol, side, price, qty,
                     tradeAmt, newPos, newAvg, floatPnl]);
  });

  ledgerSheet.clearContents();
  ledgerSheet.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]);
  if (ledgerRows.length) {
    ledgerSheet.getRange(2, 1, ledgerRows.length, LEDGER_HEADERS.length)
               .setValues(ledgerRows);
  }
}

/** Triggered when a user edits the spreadsheet */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === DATA_SHEET_NAME && e.range.getRow() >= TRADE_HEADER_ROW) {
    syncLedgerWithData();
  }
}

/** Triggered on structural changes (row deletion etc.) */
function onChange(e) {
  if (!e) return;
  if (['REMOVE_ROW','INSERT_ROW','INSERT_COLUMN','REMOVE_COLUMN'].indexOf(e.changeType) >= 0) {
    syncLedgerWithData();
  }
}

/** Initialise sheets when the spreadsheet is opened */
function onOpen() {
  ensureTradeArea();
  ensureLedgerSheet();
  syncLedgerWithData();
}

/** Manually rebuild the ledger */
function rebuildLedger() { syncLedgerWithData(); }
