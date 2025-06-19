/**
 * Crypto Trading Sheet Manager
 *
 * Maintains a price log, manual trade log and an automated ledger
 * in a Google Spreadsheet.
 *
 * This script creates the required "Data" and "Ledger" sheets if they
 * do not exist and keeps the ledger in sync with manual trades.
 */

const DATA_SHEET_NAME = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';

const PRICE_HEADERS  = ['Timestamp', 'BTC', 'ETH', 'SOL'];
const TRADE_HEADERS  = ['Symbol', 'Side', 'Quantity', 'Price', 'Trade Time', 'Note'];
const LEDGER_HEADERS = [
  'Trade ID', 'Trade Time', 'Symbol', 'Side', 'Price', 'Quantity',
  'Trade Amount', 'Running Position', 'Average Cost', 'Floating P&L'
];

// Initial row for the trade log header when creating the Data sheet
const INITIAL_TRADE_HEADER_ROW = 50;

// -----------------------------------------------------------------------------
// Initialization and menu
// -----------------------------------------------------------------------------

/** Adds a custom menu and ensures sheets exist */
function onOpen() {
  initialize();
  SpreadsheetApp.getUi().createMenu('Crypto Tools')
    .addItem('Fetch Latest Prices', 'menuFetchPrices')
    .addItem('Add Manual Trade', 'menuAddTrade')
    .addItem('Rebuild Ledger', 'menuRebuildLedger')
    .addToUi();
}

/** Ensure Data and Ledger sheets exist and have valid headers */
function initialize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let data   = ss.getSheetByName(DATA_SHEET_NAME);
  let ledger = ss.getSheetByName(LEDGER_SHEET_NAME);

  if (!data) {
    data = ss.insertSheet(DATA_SHEET_NAME);
    setupDataSheet(data);
  } else if (!checkDataSheet(data)) {
    SpreadsheetApp.getUi().alert('Data sheet structure invalid. Rebuilding.');
    ss.deleteSheet(data);
    data = ss.insertSheet(DATA_SHEET_NAME);
    setupDataSheet(data);
  }

  if (!ledger) {
    ledger = ss.insertSheet(LEDGER_SHEET_NAME);
    setupLedgerSheet(ledger);
  } else if (!checkLedgerSheet(ledger)) {
    SpreadsheetApp.getUi().alert('Ledger sheet structure invalid. Rebuilding.');
    ss.deleteSheet(ledger);
    ledger = ss.insertSheet(LEDGER_SHEET_NAME);
    setupLedgerSheet(ledger);
  }
}

/** Build the Data sheet headers and spacing */
function setupDataSheet(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, PRICE_HEADERS.length)
       .setValues([PRICE_HEADERS])
       .setFontWeight('bold');
  sheet.getRange(INITIAL_TRADE_HEADER_ROW, 1, 1, TRADE_HEADERS.length)
       .setValues([TRADE_HEADERS])
       .setFontWeight('bold');
}

/** Build the Ledger sheet headers */
function setupLedgerSheet(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, LEDGER_HEADERS.length)
       .setValues([LEDGER_HEADERS])
       .setFontWeight('bold');
}

/** Validate Data sheet header rows */
function checkDataSheet(sheet) {
  if (!sheet) return false;
  const hdr = sheet.getRange(1, 1, 1, PRICE_HEADERS.length).getValues()[0];
  for (let i = 0; i < PRICE_HEADERS.length; i++) {
    if (hdr[i] !== PRICE_HEADERS[i]) return false;
  }
  const thr = getTradeHeaderRow(sheet);
  if (!thr) return false;
  const tHdr = sheet.getRange(thr, 1, 1, TRADE_HEADERS.length).getValues()[0];
  for (let i = 0; i < TRADE_HEADERS.length; i++) {
    if (tHdr[i] !== TRADE_HEADERS[i]) return false;
  }
  return true;
}

/** Validate Ledger sheet header */
function checkLedgerSheet(sheet) {
  if (!sheet) return false;
  const hdr = sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).getValues()[0];
  for (let i = 0; i < LEDGER_HEADERS.length; i++) {
    if (hdr[i] !== LEDGER_HEADERS[i]) return false;
  }
  return true;
}

// -----------------------------------------------------------------------------
// Menu actions
// -----------------------------------------------------------------------------

/** Menu: fetch latest prices and append to Data sheet */
function menuFetchPrices() {
  initialize();
  appendLatestPrices();
}

/** Menu: prompt user and append a trade */
function menuAddTrade() {
  initialize();
  const ui = SpreadsheetApp.getUi();

  const symResp = ui.prompt('Add Trade', 'Symbol (e.g. BTC)', ui.ButtonSet.OK_CANCEL);
  if (symResp.getSelectedButton() !== ui.Button.OK) return;
  const symbol = symResp.getResponseText().trim().toUpperCase();
  if (!symbol) return;

  const sideResp = ui.prompt('Add Trade', 'Side (Buy or Sell)', ui.ButtonSet.OK_CANCEL);
  if (sideResp.getSelectedButton() !== ui.Button.OK) return;
  const side = sideResp.getResponseText().trim();
  if (!/^buy$|^sell$/i.test(side)) { ui.alert('Side must be Buy or Sell.'); return; }

  const qtyResp = ui.prompt('Add Trade', 'Quantity', ui.ButtonSet.OK_CANCEL);
  if (qtyResp.getSelectedButton() !== ui.Button.OK) return;
  const qty = parseFloat(qtyResp.getResponseText());
  if (isNaN(qty)) { ui.alert('Invalid quantity.'); return; }

  const priceResp = ui.prompt('Add Trade', 'Price', ui.ButtonSet.OK_CANCEL);
  if (priceResp.getSelectedButton() !== ui.Button.OK) return;
  const price = parseFloat(priceResp.getResponseText());
  if (isNaN(price)) { ui.alert('Invalid price.'); return; }

  const noteResp = ui.prompt('Add Trade', 'Note (optional)', ui.ButtonSet.OK_CANCEL);
  if (noteResp.getSelectedButton() !== ui.Button.OK) return;
  const note = noteResp.getResponseText();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  appendTradeRow(sheet, [symbol, side, qty, price, new Date(), note]);
  buildLedger();
}

/** Menu: rebuild the ledger from trade log */
function menuRebuildLedger() {
  initialize();
  buildLedger();
}

// -----------------------------------------------------------------------------
// Price fetching and logging
// -----------------------------------------------------------------------------

/** Fetch spot price for a symbol using Coinbase API */
function fetchPrice(symbol) {
  const url = 'https://api.coinbase.com/v2/prices/' + symbol + '-USD/spot';
  try {
    const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    if (res.getResponseCode() === 200) {
      const json = JSON.parse(res.getContentText());
      return parseFloat(json.data.amount);
    }
  } catch (err) {}
  return null;
}

/** Fetch the latest BTC, ETH and SOL prices */
function fetchLatestPrices() {
  return {
    BTC: fetchPrice('BTC'),
    ETH: fetchPrice('ETH'),
    SOL: fetchPrice('SOL')
  };
}

/** Append latest prices to the Data sheet */
function appendLatestPrices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  if (!checkDataSheet(sheet)) {
    SpreadsheetApp.getUi().alert('Data sheet structure invalid.');
    return;
  }

  let tradeHeaderRow = getTradeHeaderRow(sheet);
  if (!tradeHeaderRow) tradeHeaderRow = createTradeHeader(sheet);

  const prices = fetchLatestPrices();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  let nextRow = findNextPriceRow(sheet, tradeHeaderRow);
  if (nextRow >= tradeHeaderRow) {
    sheet.insertRowsBefore(tradeHeaderRow, 1);
    tradeHeaderRow += 1;
    nextRow = tradeHeaderRow - 1;
  }

  sheet.getRange(nextRow, 1, 1, PRICE_HEADERS.length)
       .setValues([[ts, prices.BTC, prices.ETH, prices.SOL]]);
}

// -----------------------------------------------------------------------------
// Trade log helpers
// -----------------------------------------------------------------------------

/** Append a trade row to the manual trade log */
function appendTradeRow(sheet, values) {
  let thr = getTradeHeaderRow(sheet);
  if (!thr) thr = createTradeHeader(sheet);
  let row = thr + 1;
  while (sheet.getRange(row, 1, 1, TRADE_HEADERS.length).getValues()[0].some(v => v !== '')) row++;
  sheet.getRange(row, 1, 1, TRADE_HEADERS.length).setValues([values]);
}

/** Locate the trade log header row */
function getTradeHeaderRow(sheet) {
  const last = sheet.getLastRow();
  for (let r = 1; r <= last; r++) {
    const row = sheet.getRange(r, 1, 1, TRADE_HEADERS.length).getValues()[0];
    let match = true;
    for (let i = 0; i < TRADE_HEADERS.length; i++) {
      if (row[i] !== TRADE_HEADERS[i]) { match = false; break; }
    }
    if (match) return r;
  }
  return null;
}

/** Create the trade log header below existing data */
function createTradeHeader(sheet) {
  const row = sheet.getLastRow() + 2;
  sheet.getRange(row, 1, 1, TRADE_HEADERS.length)
       .setValues([TRADE_HEADERS])
       .setFontWeight('bold');
  return row;
}

/** Find the next price row before the trade log */
function findNextPriceRow(sheet, tradeHeaderRow) {
  const rng = sheet.getRange(2, 1, tradeHeaderRow - 2, 1).getValues();
  for (let i = rng.length - 1; i >= 0; i--) {
    if (rng[i][0]) return i + 3;
  }
  return 2;
}

// -----------------------------------------------------------------------------
// Ledger building
// -----------------------------------------------------------------------------

/** Recalculate the entire ledger based on the trade log */
function buildLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data   = ss.getSheetByName(DATA_SHEET_NAME);
  const ledger = ss.getSheetByName(LEDGER_SHEET_NAME);

  if (!checkDataSheet(data) || !checkLedgerSheet(ledger)) {
    SpreadsheetApp.getUi().alert('Sheet structure invalid. Cannot rebuild ledger.');
    return;
  }

  const thr = getTradeHeaderRow(data);
  if (!thr) {
    SpreadsheetApp.getUi().alert('Trade log header missing.');
    return;
  }

  const lastPriceRow = findNextPriceRow(data, thr) - 1;
  const priceRow = data.getRange(lastPriceRow, 1, 1, PRICE_HEADERS.length).getValues()[0];
  const currentPrices = {BTC: priceRow[1], ETH: priceRow[2], SOL: priceRow[3]};

  let trades = [];
  if (data.getLastRow() > thr) {
    trades = data.getRange(thr + 1, 1, data.getLastRow() - thr, TRADE_HEADERS.length)
                .getValues()
                .filter(r => r.some(v => v !== ''));
  }

  const pos = {};    // running position per symbol
  const avg = {};    // average cost per symbol
  const rows = [];
  let id = 1;

  trades.forEach(t => {
    const [symbol, side, qtyVal, priceVal, timeVal] = t;
    const qty = parseFloat(qtyVal);
    const price = parseFloat(priceVal);
    if (!symbol || isNaN(qty) || isNaN(price) || !side) return;

    const sign = /^buy$/i.test(side) ? 1 : -1;
    if (pos[symbol] === undefined) { pos[symbol] = 0; avg[symbol] = 0; }

    const prevPos = pos[symbol];
    const prevAvg = avg[symbol];
    const newPos = prevPos + qty * sign;
    let newAvg = prevAvg;

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
    const floatPnl = ((currentPrices[symbol] || 0) - newAvg) * newPos;

    rows.push([id++, timeVal || new Date(), symbol, side, price, qty,
               tradeAmt, newPos, newAvg, floatPnl]);
  });

  ledger.clearContents();
  ledger.getRange(1, 1, 1, LEDGER_HEADERS.length)
        .setValues([LEDGER_HEADERS])
        .setFontWeight('bold');
  if (rows.length) {
    ledger.getRange(2, 1, rows.length, LEDGER_HEADERS.length).setValues(rows);
  }
}

// -----------------------------------------------------------------------------
// Triggered updates
// -----------------------------------------------------------------------------

/** When the Data sheet is edited, rebuild the ledger if trade log changed */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === DATA_SHEET_NAME) {
    const thr = getTradeHeaderRow(sheet);
    if (thr && e.range.getRow() > thr) {
      buildLedger();
    }
  }
}
