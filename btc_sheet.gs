// Google Apps Script integrating Binance to Google Sheets add-on
// and providing a simple two-sheet trading workflow.

/**
 * Creates the "Data" and "ledger" sheets, inserts headers,
 * adds a 2‑hour trigger for updateData() and runs it once.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Setup Data sheet
  var dataSheet = ss.getSheetByName('Data');
  if (!dataSheet) {
    dataSheet = ss.insertSheet('Data');
  }
  var dataHeaders = [['Timestamp', 'BTC', 'ETH', 'SOL']];
  dataSheet.clear();
  dataSheet.getRange(1, 1, 1, dataHeaders[0].length)
    .setValues(dataHeaders)
    .setFontWeight('bold');

  // Setup ledger sheet
  var ledgerSheet = ss.getSheetByName('ledger');
  if (!ledgerSheet) {
    ledgerSheet = ss.insertSheet('ledger');
  }
  var ledgerHeaders = [['Timestamp', 'Asset', 'Action', 'Quantity', 'Price', 'Running P&L']];
  ledgerSheet.clear();
  ledgerSheet.getRange(1, 1, 1, ledgerHeaders[0].length)
    .setValues(ledgerHeaders)
    .setFontWeight('bold');

  // Create time-driven trigger if not already present
  var exists = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'updateData';
  });
  if (!exists) {
    ScriptApp.newTrigger('updateData').timeBased().everyHours(2).create();
  }

  // Fetch initial data
  updateData();
}

/**
 * Fetches last 24h of 2h OHLCV data from Binance for BTC, ETH and SOL,
 * writes it to the Data sheet and recalculates P&L.
 */
function updateData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  if (!sheet) {
    SpreadsheetApp.getActive().toast('Data sheet missing - run setup()', 'Error', 5);
    return;
  }

  var symbols = ['BTCUSDT', 'ETHUSDT', 'SOLUSDT'];
  var rows = [];

  try {
    symbols.forEach(function(sym, idx) {
      var data = BINANCE('history', sym, 'interval: 2h, limit: 12');
      if (!data || !data.length) throw new Error('No data for ' + sym);

      data.forEach(function(rec, i) {
        var ts = new Date(rec[0]);
        var iso = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
        if (!rows[i]) rows[i] = [iso, '', '', ''];
        rows[i][idx + 1] = rec[4]; // close price
      });

      Utilities.sleep(1100); // avoid rate limit
    });

    sheet.getRange(2, 1, 999, 4).clearContent();
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }
  } catch (e) {
    Logger.log('updateData error: ' + e.toString());
    SpreadsheetApp.getActive().toast('updateData error: ' + e.message, 'Error', 5);
  }

  calculatePNL();
}

/**
 * Calculates running P&L for trades recorded in the ledger using FIFO.
 */
function calculatePNL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ledger = ss.getSheetByName('ledger');
  if (!ledger || ledger.getLastRow() < 2) return;

  var dataSheet = ss.getSheetByName('Data');
  var prices = { BTC: 0, ETH: 0, SOL: 0 };
  if (dataSheet && dataSheet.getLastRow() >= 2) {
    var p = dataSheet.getRange(dataSheet.getLastRow(), 2, 1, 3).getValues()[0];
    prices = { BTC: parseFloat(p[0]) || 0, ETH: parseFloat(p[1]) || 0, SOL: parseFloat(p[2]) || 0 };
  }

  var entries = ledger.getRange(2, 1, ledger.getLastRow() - 1, 5).getValues();
  var lots = { BTC: [], ETH: [], SOL: [] };
  var realized = { BTC: 0, ETH: 0, SOL: 0 };
  var results = [];
  var totalCost = 0;

  entries.forEach(function(r) {
    var asset = (r[1] || '').toUpperCase();
    var action = (r[2] || '').toLowerCase();
    var qty = parseFloat(r[3]) || 0;
    var price = parseFloat(r[4]) || 0;
    if (!asset || !qty || !price) {
      results.push(['']);
      return;
    }

    if (action === 'buy') {
      lots[asset].push({ qty: qty, price: price });
      totalCost += qty * price;
    } else if (action === 'sell') {
      var remaining = qty;
      var cost = 0;
      while (remaining > 0 && lots[asset].length) {
        var lot = lots[asset][0];
        var take = Math.min(lot.qty, remaining);
        cost += take * lot.price;
        lot.qty -= take;
        if (lot.qty <= 0) lots[asset].shift();
        remaining -= take;
      }
      realized[asset] += qty * price - cost;
    }

    var unrealized = 0;
    Object.keys(lots).forEach(function(key) {
      var qtyHeld = 0;
      var costHeld = 0;
      lots[key].forEach(function(l) {
        qtyHeld += l.qty;
        costHeld += l.qty * l.price;
      });
      if (qtyHeld) {
        unrealized += qtyHeld * (prices[key] || 0) - costHeld;
      }
    });

    var totalReal = realized.BTC + realized.ETH + realized.SOL;
    results.push([totalReal + unrealized]);
  });

  if (results.length) {
    ledger.getRange(2, 6, results.length, 1).setValues(results);
  }

  var finalPnL = results.length ? results[results.length - 1][0] : 0;
  var returnPct = totalCost ? (finalPnL / totalCost) * 100 : 0;
  ledger.getRange(results.length + 3, 5, 1, 2)
    .setValues([[ 'Total Return %', returnPct ]]);

  ledger.getRange(1, 1, 1, 6).setFontWeight('bold');
  ledger.autoResizeColumns(1, 6);
}

/**
 * Adds a custom menu prompting for authorization if needed.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Crypto Sheets');
  var needAuth = false;
  try {
    BINANCE('version');
  } catch (e) {
    needAuth = true;
  }
  if (needAuth) {
    menu.addItem('Authorize add-on!', 'authorizeAddon');
  }
  menu.addToUi();
}

/**
 * Simple wrapper to trigger the Binance authorization flow.
 */
function authorizeAddon() {
  try {
    BINANCE('ping');
  } catch (e) {
    SpreadsheetApp.getActive().toast('Authorization required: ' + e.message, 'BINANCE', 5);
  }
}

/**
 * Recalculates P&L when the ledger sheet is edited.
 */
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === 'ledger') {
    calculatePNL();
  }
}

// Run setup() once manually, then click Crypto Sheets → Authorize add-on! to grant permissions.
