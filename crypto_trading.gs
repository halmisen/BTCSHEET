/**
 * Crypto Tools for Google Sheets
 *
 * Creates a custom menu that allows fetching the latest BTC/ETH/SOL prices,
 * manually logging trades and rebuilding the trade ledger.
 * All sheets are created automatically if missing.
 */

const DATA_SHEET_NAME = 'Data';
const LEDGER_SHEET_NAME = 'Ledger';

const PRICE_HEADERS = ['Timestamp', 'BTC', 'ETH', 'SOL'];
const LEDGER_HEADERS = [
  'Trade ID', 'Trade Time', 'Symbol', 'Side', 'Price', 'Quantity',
  'Trade Amount', 'Running Position', 'Average Cost', 'Floating P&L'
];

/** Adds the custom menu when the spreadsheet is opened. */
function onOpen() {
  ensureSheets();
  SpreadsheetApp.getUi()
    .createMenu('Crypto Tools')
    .addItem('Fetch Latest Prices', 'menuFetchLatestPrices')
    .addItem('Add Manual Trade', 'menuAddManualTrade')
    .addItem('Rebuild Ledger', 'menuRebuildLedger')
    .addToUi();
}

/** Ensures that Data and Ledger sheets exist with proper headers. */
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let data = ss.getSheetByName(DATA_SHEET_NAME);
  if (!data) {
    data = ss.insertSheet(DATA_SHEET_NAME);
  }
  const dataHdr = data.getRange(1, 1, 1, PRICE_HEADERS.length).getValues()[0];
  if (dataHdr.join() !== PRICE_HEADERS.join()) {
    data.clear();
    data.getRange(1, 1, 1, PRICE_HEADERS.length)
        .setValues([PRICE_HEADERS])
        .setFontWeight('bold');
  }

  let ledger = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (!ledger) {
    ledger = ss.insertSheet(LEDGER_SHEET_NAME);
  }
  const ledgerHdr = ledger.getRange(1, 1, 1, LEDGER_HEADERS.length).getValues()[0];
  if (ledgerHdr.join() !== LEDGER_HEADERS.join()) {
    ledger.clear();
    ledger.getRange(1, 1, 1, LEDGER_HEADERS.length)
          .setValues([LEDGER_HEADERS])
          .setFontWeight('bold');
  }
}

/** Menu handler: fetches prices and appends a new row to the Data sheet. */
function menuFetchLatestPrices() {
  ensureSheets();
  appendLatestPrices();
}

/** Menu handler: prompts the user for trade details and saves the trade. */
function menuAddManualTrade() {
  ensureSheets();
  const ui = SpreadsheetApp.getUi();

  const symRes = ui.prompt('Add Trade', 'Symbol (e.g. BTC)', ui.ButtonSet.OK_CANCEL);
  if (symRes.getSelectedButton() !== ui.Button.OK) return;
  const symbol = symRes.getResponseText().trim().toUpperCase();
  if (!symbol) return;

  const sideRes = ui.prompt('Add Trade', 'Side (Buy or Sell)', ui.ButtonSet.OK_CANCEL);
  if (sideRes.getSelectedButton() !== ui.Button.OK) return;
  const side = sideRes.getResponseText().trim().toUpperCase();
  if (side !== 'BUY' && side !== 'SELL') { ui.alert('Side must be Buy or Sell'); return; }

  const priceRes = ui.prompt('Add Trade', 'Price', ui.ButtonSet.OK_CANCEL);
  if (priceRes.getSelectedButton() !== ui.Button.OK) return;
  const price = parseFloat(priceRes.getResponseText());
  if (isNaN(price)) { ui.alert('Invalid price'); return; }

  const qtyRes = ui.prompt('Add Trade', 'Quantity', ui.ButtonSet.OK_CANCEL);
  if (qtyRes.getSelectedButton() !== ui.Button.OK) return;
  const qty = parseFloat(qtyRes.getResponseText());
  if (isNaN(qty)) { ui.alert('Invalid quantity'); return; }

  const ledger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEDGER_SHEET_NAME);
  const last = ledger.getLastRow();
  const nextId = last >= 2 ? Number(ledger.getRange(last, 1).getValue()) + 1 : 1;
  const time = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  ledger.appendRow([nextId, time, symbol, side, price, qty, '', '', '', '']);
  menuRebuildLedger();
}

/** Menu handler: recomputes all ledger statistics. */
function menuRebuildLedger() {
  ensureSheets();
  rebuildLedger();
}

/** Fetches the latest spot price for a single symbol. */
function fetchPrice(symbol) {
  const url = 'https://api.coinbase.com/v2/prices/' + symbol + '-USD/spot';
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const json = JSON.parse(res.getContentText());
      return parseFloat(json.data.amount);
    }
    Logger.log('Failed to fetch ' + symbol + ': HTTP ' + res.getResponseCode());
  } catch (err) {
    Logger.log('Error fetching ' + symbol + ': ' + err);
  }
  return 'ERROR';
}

/** Fetches prices for BTC, ETH and SOL and appends them to the Data sheet. */
function appendLatestPrices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const ts = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const prices = [fetchPrice('BTC'), fetchPrice('ETH'), fetchPrice('SOL')];
  sheet.appendRow([ts].concat(prices));
}

/** Rebuilds the ledger based on the raw trades in the Ledger sheet. */
function rebuildLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = ss.getSheetByName(LEDGER_SHEET_NAME);
  const data = ss.getSheetByName(DATA_SHEET_NAME);

  const raw = ledger.getRange(2, 1, Math.max(ledger.getLastRow() - 1, 0), 6).getValues()
                .filter(r => r[0]);
  const results = [];
  const pos = {};
  const avg = {};

  raw.forEach(r => {
    const [id, time, sym, side, price, qty] = r;
    const sign = side.toUpperCase() === 'BUY' ? 1 : -1;
    if (pos[sym] === undefined) { pos[sym] = 0; avg[sym] = 0; }

    const prevPos = pos[sym];
    const prevAvg = avg[sym];
    const newPos = prevPos + qty * sign;
    let newAvg = prevAvg;

    if (sign > 0) {
      newAvg = (prevAvg * prevPos + price * qty) / newPos;
    } else if (newPos === 0) {
      newAvg = 0;
    }

    pos[sym] = newPos;
    avg[sym] = newAvg;

    const tradeAmt = price * qty * sign;
    results.push([id, time, sym, side, price, qty, tradeAmt, newPos, newAvg, '']);
  });

  const priceRow = data.getRange(data.getLastRow(), 1, 1, PRICE_HEADERS.length).getValues()[0];
  const currentPrices = { BTC: priceRow[1], ETH: priceRow[2], SOL: priceRow[3] };

  results.forEach(row => {
    const sym = row[2];
    const cur = currentPrices[sym];
    const posVal = row[7];
    const avgVal = row[8];
    if (typeof cur === 'number') {
      row[9] = (cur - avgVal) * posVal;
    } else {
      row[9] = 'ERROR';
    }
  });

  ledger.getRange(2, 1, Math.max(ledger.getLastRow() - 1, 0), LEDGER_HEADERS.length).clearContent();
  if (results.length) {
    ledger.getRange(2, 1, results.length, LEDGER_HEADERS.length).setValues(results);
  }
}
