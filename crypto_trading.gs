/**
 * Crypto Trading Toolkit
 *
 * Provides a custom menu with utilities to fetch cryptocurrency prices,
 * record manual trades and maintain a running trade ledger.
 *
 * The script automatically creates the required "Data" and "Ledger" sheets
 * if they do not exist. All functions include basic error handling so the
 * tools work even on a blank spreadsheet.
 */

// -----------------------------------------------------------------------------
// Configuration
// -----------------------------------------------------------------------------

/** Names of the sheets used by this script */
const DATA_SHEET_NAME = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';

/** Column headers for each sheet */
const DATA_HEADERS = ['Timestamp', 'BTC', 'ETH', 'SOL'];
const LEDGER_HEADERS = [
  'Trade ID', 'Trade Time', 'Symbol', 'Side', 'Price', 'Quantity',
  'Trade Amount', 'Running Position', 'Average Cost', 'Floating P&L'
];

// -----------------------------------------------------------------------------
// Menu setup and initialization
// -----------------------------------------------------------------------------

/**
 * Adds the "Crypto Tools" menu and ensures sheets exist when the
 * spreadsheet is opened.
 */
function onOpen() {
  ensureSheets();
  SpreadsheetApp.getUi()
    .createMenu('Crypto Tools')
    .addItem('Fetch Latest Prices', 'fetchLatestPrices')
    .addItem('Add Manual Trade', 'addManualTrade')
    .addItem('Rebuild Ledger', 'rebuildLedger')
    .addToUi();
}

/**
 * Creates the Data and Ledger sheets if they are missing and ensures the
 * correct header rows are present.
 */
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let data = ss.getSheetByName(DATA_SHEET_NAME);
  if (!data) {
    data = ss.insertSheet(DATA_SHEET_NAME);
  }
  const dataHeader = data.getRange(1, 1, 1, DATA_HEADERS.length).getValues()[0];
  if (dataHeader.join() !== DATA_HEADERS.join()) {
    data.clear();
    data.getRange(1, 1, 1, DATA_HEADERS.length)
        .setValues([DATA_HEADERS])
        .setFontWeight('bold');
  }

  let ledger = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (!ledger) {
    ledger = ss.insertSheet(LEDGER_SHEET_NAME);
  }
  const ledgerHeader = ledger.getRange(1, 1, 1, LEDGER_HEADERS.length).getValues()[0];
  if (ledgerHeader.join() !== LEDGER_HEADERS.join()) {
    ledger.clear();
    ledger.getRange(1, 1, 1, LEDGER_HEADERS.length)
          .setValues([LEDGER_HEADERS])
          .setFontWeight('bold');
  }
}

// -----------------------------------------------------------------------------
// Price fetching utilities
// -----------------------------------------------------------------------------

/**
 * Fetches the current USD spot price for a given symbol from Coinbase.
 * Returns the numeric price or the string "ERROR" on failure.
 */
function fetchPrice(symbol) {
  const url = `https://api.coinbase.com/v2/prices/${symbol}-USD/spot`;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const json = JSON.parse(res.getContentText());
      const amount = parseFloat(json.data.amount);
      if (!isNaN(amount)) return amount;
    }
    Logger.log(`Error fetching ${symbol}: HTTP ${res.getResponseCode()}`);
  } catch (err) {
    Logger.log(`Fetch failed for ${symbol}: ${err}`);
  }
  return 'ERROR';
}

/**
 * Fetches latest prices for BTC, ETH and SOL and appends a row to the
 * Data sheet with the current timestamp.
 */
function fetchLatestPrices() {
  ensureSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  const btc = fetchPrice('BTC');
  const eth = fetchPrice('ETH');
  const sol = fetchPrice('SOL');

  sheet.appendRow([ts, btc, eth, sol]);
}

// -----------------------------------------------------------------------------
// Manual trade entry
// -----------------------------------------------------------------------------

/**
 * Prompts the user for trade details and appends the trade to the Ledger
 * sheet. After inserting the trade, the entire ledger is rebuilt so all
 * statistics remain accurate.
 */
function addManualTrade() {
  ensureSheets();
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

  const ledger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEDGER_SHEET_NAME);
  const id = Math.max(ledger.getLastRow(), 1);
  const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  ledger.appendRow([id, time, symbol, side, price, qty]);
  rebuildLedger();
}

// -----------------------------------------------------------------------------
// Ledger computation
// -----------------------------------------------------------------------------

/**
 * Rebuilds the entire ledger based on the raw trades recorded in the first
 * six columns of the Ledger sheet. Running position, average cost and
 * floating P&L are recalculated from scratch for each row.
 */
function rebuildLedger() {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  const ledger = ss.getSheetByName(LEDGER_SHEET_NAME);

  let currentPrices = { BTC: 0, ETH: 0, SOL: 0 };
  if (dataSheet.getLastRow() > 1) {
    const row = dataSheet.getRange(dataSheet.getLastRow(), 1, 1, DATA_HEADERS.length).getValues()[0];
    currentPrices = { BTC: row[1], ETH: row[2], SOL: row[3] };
  }

  const trades = ledger.getLastRow() > 1
    ? ledger.getRange(2, 1, ledger.getLastRow() - 1, 6).getValues()
    : [];

  const output = [];
  const pos = {};
  const avg = {};

  trades.forEach(t => {
    let [id, time, symbol, side, price, qty] = t;
    if (!symbol || !side) return;
    qty = parseFloat(qty);
    price = parseFloat(price);
    if (isNaN(qty) || isNaN(price)) return;

    const sign = /^buy$/i.test(side) ? 1 : -1;
    const prevPos = pos[symbol] || 0;
    const prevAvg = avg[symbol] || 0;
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
    const floatPnl = currentPrices[symbol]
      ? (currentPrices[symbol] - newAvg) * newPos
      : '';

    output.push([id, time, symbol, side, price, qty, tradeAmt, newPos, newAvg, floatPnl]);
  });

  ledger.clearContents();
  ledger.getRange(1, 1, 1, LEDGER_HEADERS.length)
        .setValues([LEDGER_HEADERS])
        .setFontWeight('bold');
  if (output.length) {
    ledger.getRange(2, 1, output.length, LEDGER_HEADERS.length).setValues(output);
  }
}

// -----------------------------------------------------------------------------
// End of script
// -----------------------------------------------------------------------------
