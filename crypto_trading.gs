// Google Sheets Crypto Trading Simulator - Single File Version
// This script manages price data, trade entry and a running ledger.
// It automatically creates required sheets and keeps the Ledger in sync with
// the Data sheet. Designed for easy extension when adding new tokens/fields.
// This script only alters rows at or below TRADE_HEADER_ROW so the price data region is preserved.

// --- Configuration ----------------------------------------------------------
const DATA_SHEET_NAME   = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';
const TRADE_HEADER_ROW  = 20;            // trade table header row (1-indexed)
const TEMPLATE_ROWS     = 5;             // empty rows for manual entry

const PRICE_HEADERS   = ['Symbol','Price','Δ2h %','Δ4h %','Δ12h %','Δ24h %'];
const PRODUCTS        = ['BTC-USD','ETH-USD','SOL-USD'];

const TRADE_HEADERS  = ['Symbol','Side','Quantity','Price','Trade Time','Note'];
const LEDGER_HEADERS = [
  'Trade ID','Trade Time','Symbol','Side','Price','Quantity',
  'Trade Amount','Running Position','Average Cost','Floating P&L'
];

// --- Sheet Initialisation --------------------------------------------------

/** Ensure the required Data and Ledger sheets exist.
 *  If missing, rename default 'Sheet#' sheets or create new ones.
 *  @return {{data:Sheet, ledger:Sheet}}
 */
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let data = ss.getSheetByName(DATA_SHEET_NAME);
  let ledger = ss.getSheetByName(LEDGER_SHEET_NAME);

  const unused = ss.getSheets().filter(s => /^Sheet\d+$/.test(s.getName()));

  if (!data) {
    data = unused.shift();
    data = data ? data.setName(DATA_SHEET_NAME) : ss.insertSheet(DATA_SHEET_NAME);
  }

  if (!ledger) {
    ledger = unused.shift();
    ledger = ledger ? ledger.setName(LEDGER_SHEET_NAME)
                    : ss.insertSheet(LEDGER_SHEET_NAME);
  }
  return {data: data, ledger: ledger};
}

/** Ensure trade input headers and template rows in the Data sheet */
function ensureTradeArea(sheet) {
  if (!sheet) return;
  const rng = sheet.getRange(TRADE_HEADER_ROW, 1, 1, TRADE_HEADERS.length);
  const cur = rng.getValues()[0];
  let mismatch = false;
  for (let i = 0; i < TRADE_HEADERS.length; i++) {
    if (cur[i] !== TRADE_HEADERS[i]) { mismatch = true; break; }
  }
  if (mismatch) rng.setValues([TRADE_HEADERS]);

  // Always ensure template rows exist below the header
  const tempRange = sheet.getRange(TRADE_HEADER_ROW + 1, 1,
                                  TEMPLATE_ROWS, TRADE_HEADERS.length);
  tempRange.clearContent();
}

/** Ensure the Ledger sheet headers exist and nothing else on row 1 */
function ensureLedgerHeaders(sheet) {
  if (!sheet) return;
  const rng = sheet.getRange(1, 1, 1, LEDGER_HEADERS.length);
  const cur = rng.getValues()[0];
  let mismatch = false;
  for (let i = 0; i < LEDGER_HEADERS.length; i++) {
    if (cur[i] !== LEDGER_HEADERS[i]) { mismatch = true; break; }
  }
  if (mismatch) {
    sheet.clear();
    rng.setValues([LEDGER_HEADERS]);
  }
}

// --- Helper ----------------------------------------------------------------

/** Ensure the price summary table headers exist */
function ensurePriceTable(sheet) {
  if (!sheet) return;
  const rng = sheet.getRange(1, 1, 1, PRICE_HEADERS.length);
  const cur = rng.getValues()[0];
  let mismatch = false;
  for (let i = 0; i < PRICE_HEADERS.length; i++) {
    if (cur[i] !== PRICE_HEADERS[i]) { mismatch = true; break; }
  }
  if (mismatch) rng.setValues([PRICE_HEADERS]);
}

/** Fetch the latest `count` hourly candles for a product */
function fetchHourlyCandles(product, count) {
  const url = `https://api.exchange.coinbase.com/products/${product}/candles?granularity=3600&limit=${count}`;
  const res = UrlFetchApp.fetch(url, {headers: {Accept: 'application/json'}});
  const data = JSON.parse(res.getContentText());
  data.sort((a, b) => a[0] - b[0]);
  return data;
}

/** Build a price row for the summary table */
function calcPriceRow(product) {
  const candles = fetchHourlyCandles(product, 25); // 24h of data
  const last = candles[candles.length - 1][4];
  function pct(h) {
    const i = candles.length - 1 - h;
    if (i < 0) return '';
    const p = candles[i][4];
    return (last - p) / p;
  }
  return [product.split('-')[0], last, pct(2), pct(4), pct(12), pct(24)];
}

/** Refresh the price summary table */
function refreshPriceSummary() {
  const {data} = ensureSheets();
  ensurePriceTable(data);
  const rows = PRODUCTS.map(p => calcPriceRow(p));
  data.getRange(2, 1, rows.length, PRICE_HEADERS.length).setValues(rows);
  const extra = data.getLastRow() - (rows.length + 1);
  if (extra > 0) {
    data.getRange(rows.length + 2, 1, extra, PRICE_HEADERS.length).clearContent();
  }
}

/** Ensure a 2h trigger exists for refreshing prices */
function ensurePriceTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === 'refreshPriceSummary');
  if (!exists) {
    ScriptApp.newTrigger('refreshPriceSummary').timeBased().everyHours(2).create();
  }
}

/** Read the latest price for each token column in the Data sheet */
function getLatestPrices(sheet) {
  const last = sheet.getLastRow();
  const range = last > 1 ? sheet.getRange(2, 1, last - 1, 2).getValues() : [];
  const prices = {};
  range.forEach(row => {
    const sym = row[0];
    const price = parseFloat(row[1]);
    if (sym && !isNaN(price)) prices[sym] = price;
  });
  return prices;
}

/** Rebuild the Ledger based on the trades entered in the Data sheet */
function syncLedgerWithData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const {data, ledger} = ensureSheets();
  ensurePriceTable(data);
  ensureTradeArea(data);
  ensureLedgerHeaders(ledger);

  const lastRow = data.getLastRow();
  const rows = lastRow > TRADE_HEADER_ROW
    ? data.getRange(TRADE_HEADER_ROW + 1, 1,
                    lastRow - TRADE_HEADER_ROW, TRADE_HEADERS.length).getValues()
    : [];

  const latestPrices = getLatestPrices(data);
  const position = {};   // running position per symbol
  const avgCost  = {};   // average cost per symbol
  const ledgerRows = [];

  rows.forEach(row => {
    const [symbol, side, qtyVal, priceVal, timeVal] = row;
    const qty = parseFloat(qtyVal);
    const price = parseFloat(priceVal);
    if (!symbol || !side || isNaN(qty) || isNaN(price)) return; // skip blanks
    const time = timeVal || new Date();

    if (position[symbol] === undefined) { position[symbol] = 0; avgCost[symbol] = 0; }
    const sign = side.toString().toLowerCase() === 'buy' ? 1 : -1;

    const prevPos = position[symbol];
    const prevAvg = avgCost[symbol];
    const newPos  = prevPos + qty * sign;
    let   newAvg  = prevAvg;
    if (sign > 0) {
      // buying more
      newAvg = (prevAvg * Math.abs(prevPos) + price * qty) / Math.abs(newPos);
    } else {
      // selling
      if (Math.sign(prevPos) === Math.sign(newPos) && prevPos !== 0) {
        newAvg = prevAvg;
      } else if (newPos === 0) {
        newAvg = 0;
      } else {
        newAvg = price;
      }
    }
    position[symbol] = newPos;
    avgCost[symbol] = newAvg;

    const tradeAmt = price * qty * sign;
    const floatPnl = ((latestPrices[symbol] || 0) - newAvg) * newPos;
    ledgerRows.push([ledgerRows.length + 1, time, symbol, side, price,
                     qty, tradeAmt, newPos, newAvg, floatPnl]);
  });

  ledger.clearContents();
  ledger.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]);
  if (ledgerRows.length) {
    ledger.getRange(2, 1, ledgerRows.length, LEDGER_HEADERS.length)
          .setValues(ledgerRows);
  }
}

// --- Triggers ---------------------------------------------------------------

/** Called whenever a user edits the spreadsheet */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === DATA_SHEET_NAME &&
      e.range.getRow() >= TRADE_HEADER_ROW) {
    syncLedgerWithData();
  }
}

/** Called on structural changes like row insertions/deletions */
function onChange(e) {
  if (!e) return;
  const t = e.changeType;
  if (['REMOVE_ROW','INSERT_ROW','INSERT_COLUMN','REMOVE_COLUMN'].indexOf(t) >= 0) {
    syncLedgerWithData();
  }
}

/** Initialisation when the spreadsheet is opened */
function onOpen() {
  const {data, ledger} = ensureSheets();
  ensurePriceTable(data);
  ensureTradeArea(data);
  ensureLedgerHeaders(ledger);
  refreshPriceSummary();
  ensurePriceTrigger();
  syncLedgerWithData();
}

/** Manual trigger to rebuild the ledger */
function rebuildLedger() { syncLedgerWithData(); }
