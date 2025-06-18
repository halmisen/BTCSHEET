// Google Sheets Crypto Trading Simulator - Single File Version
// This script manages price data, trade entry and a running ledger.
// It automatically creates required sheets and keeps the Ledger in sync with
// the Data sheet. Designed for easy extension when adding new tokens/fields.
// This script only alters rows at or below TRADE_HEADER_ROW so the price data region is preserved.

// --- Configuration ----------------------------------------------------------
const DATA_SHEET_NAME   = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';

// Price history layout
const PRICE_HEADERS     = ['Timestamp','BTC','ETH','SOL'];
const PRICE_TABLE_ROWS  = 50;            // rows reserved for history (incl. header)
const SUMMARY_START_ROW = PRICE_TABLE_ROWS + 1; // row after history section
const SUMMARY_ROWS      = 4;             // header + one row per coin
const TRADE_HEADER_ROW  = SUMMARY_START_ROW + SUMMARY_ROWS; // trade table header
const TEMPLATE_ROWS     = 5;             // empty rows for manual entry

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

/** Ensure the price history and summary tables exist */
function ensurePriceSection(sheet) {
  if (!sheet) return;
  const headerRng = sheet.getRange(1, 1, 1, PRICE_HEADERS.length);
  const cur = headerRng.getValues()[0];
  let mismatch = false;
  for (let i = 0; i < PRICE_HEADERS.length; i++) {
    if (cur[i] !== PRICE_HEADERS[i]) { mismatch = true; break; }
  }
  if (mismatch) headerRng.setValues([PRICE_HEADERS]);

  // ensure fixed size for history region
  const body = sheet.getRange(2, 1, PRICE_TABLE_ROWS - 1, PRICE_HEADERS.length);
  if (body.getValues().length !== PRICE_TABLE_ROWS - 1) {
    body.clearContent();
    body.setValues(Array(PRICE_TABLE_ROWS - 1)
      .fill(0)
      .map(() => Array(PRICE_HEADERS.length).fill('')));
  }

  ensureSummaryTable(sheet);
}

/** Ensure summary table header exists */
function ensureSummaryTable(sheet) {
  const hdr = ['Coin','Δ2h','Δ4h','Δ12h','Δ24h'];
  sheet.getRange(SUMMARY_START_ROW, 1, 1, hdr.length).setValues([hdr]);
  // clear coin rows if missing
  const coinRows = sheet.getRange(SUMMARY_START_ROW + 1, 1,
                                  SUMMARY_ROWS - 1, hdr.length);
  const empty = Array(SUMMARY_ROWS - 1).fill(0)
                .map(() => Array(hdr.length).fill(''));
  coinRows.setValues(empty);
}

// --- Helper ----------------------------------------------------------------

/** Read the latest price for each token column in the Data sheet */
function getLatestPrices(sheet) {
  const body = sheet.getRange(2, 1, PRICE_TABLE_ROWS - 1, PRICE_HEADERS.length).getValues();
  let last = null;
  for (let i = body.length - 1; i >= 0; i--) {
    if (body[i][0]) { last = body[i]; break; }
  }
  const prices = {};
  if (last) {
    PRICE_HEADERS.slice(1).forEach((h, idx) => { prices[h] = last[idx + 1]; });
  }
  return prices;
}

/** Fetch a single spot price from Coinbase */
function fetchSpotPrice(symbol) {
  const url = `https://api.coinbase.com/v2/prices/${symbol}-USD/spot`;
  try {
    const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    if (res.getResponseCode() !== 200) return null;
    const json = JSON.parse(res.getContentText());
    return parseFloat(json.data.amount);
  } catch (err) {
    Logger.log('Error fetching ' + symbol + ': ' + err);
    return null;
  }
}

/** Fetch all latest prices */
function fetchLatestSpotPrices() {
  return {
    BTC: fetchSpotPrice('BTC'),
    ETH: fetchSpotPrice('ETH'),
    SOL: fetchSpotPrice('SOL')
  };
}

/** Append a new price row into the fixed history region */
function appendPriceRow(sheet, prices) {
  const bodyRange = sheet.getRange(2, 1, PRICE_TABLE_ROWS - 1, PRICE_HEADERS.length);
  let data = bodyRange.getValues();
  // Remove completely empty top rows
  while (data.length && data[0].every(v => v === '')) data.shift();
  // Trim to existing data
  data = data.filter(r => r.some(v => v !== ''));
  data.push([formatTimestamp(new Date()), prices.BTC, prices.ETH, prices.SOL]);
  while (data.length > PRICE_TABLE_ROWS - 1) data.shift();
  while (data.length < PRICE_TABLE_ROWS - 1) data.unshift(['','','','']);
  bodyRange.setValues(data);
}

/** Format date as string */
function formatTimestamp(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

/** Build summary table below the history */
function buildSummary(sheet) {
  const body = sheet.getRange(2, 1, PRICE_TABLE_ROWS - 1, PRICE_HEADERS.length).getValues()
                  .filter(r => r[0]);
  if (!body.length) return;

  const latest = body[body.length - 1];
  const offsets = [1,2,6,12];
  const headers = ['Coin','Δ2h','Δ4h','Δ12h','Δ24h'];
  const coins = ['BTC','ETH','SOL'];
  const values = [headers];

  coins.forEach((c, idx) => {
    const current = latest[idx + 1];
    const row = [c];
    offsets.forEach(off => {
      const i = body.length - 1 - off;
      if (i >= 0) {
        const prev = body[i][idx + 1];
        const diff = current - prev;
        const pct = prev ? diff / prev * 100 : 0;
        const sign = diff >= 0 ? '+' : '';
        row.push(`${sign}${pct.toFixed(2)}% (${sign}${diff.toFixed(2)})`);
      } else {
        row.push('');
      }
    });
    values.push(row);
  });

  const rng = sheet.getRange(SUMMARY_START_ROW, 1, values.length, values[0].length);
  rng.clearContent();
  rng.setValues(values);
}

/** Fetch prices and update sheet */
function updatePrices() {
  const {data} = ensureSheets();
  ensurePriceSection(data);
  const prices = fetchLatestSpotPrices();
  appendPriceRow(data, prices);
  buildSummary(data);
}

/** Rebuild the Ledger based on the trades entered in the Data sheet */
function syncLedgerWithData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const {data, ledger} = ensureSheets();
  ensurePriceSection(data);
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
  ensurePriceSection(data);
  ensureTradeArea(data);
  ensureLedgerHeaders(ledger);
  buildSummary(data);
  syncLedgerWithData();
}

/** Manual trigger to rebuild the ledger */
function rebuildLedger() { syncLedgerWithData(); }

/**
 * Append the latest BTC, ETH and SOL spot prices to the "Data" sheet.
 * Fetches prices from Coinbase and writes them with the current timestamp
 * in yyyy-MM-dd HH:mm:ss format. Errors are written as "ERROR".
 */
function appendLatestPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    const msg = 'Sheet "' + DATA_SHEET_NAME + '" not found.' +
                ' Please create a sheet named "' + DATA_SHEET_NAME + '".';
    SpreadsheetApp.getUi().alert(msg);
    Logger.log(msg);
    return;
  }

  const prices = fetchLatestSpotPrices();
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues([[
    timestamp,
    prices.BTC !== null ? prices.BTC : 'ERROR',
    prices.ETH !== null ? prices.ETH : 'ERROR',
    prices.SOL !== null ? prices.SOL : 'ERROR'
  ]]);
}
