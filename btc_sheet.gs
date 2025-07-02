// Google Apps Script to fetch crypto prices and manage a simple trade ledger

/**
 * Initial setup: creates sheets and a time-driven trigger.
 * Run this once manually.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure "data" sheet exists with headers
  var dataSheet = ss.getSheetByName('data');
  if (!dataSheet) {
    dataSheet = ss.insertSheet('data');
    dataSheet.appendRow(['Timestamp', 'BTC', 'ETH', 'SOL']);
  }

  // Ensure "ledger" sheet exists with headers
  var ledgerSheet = ss.getSheetByName('ledger');
  if (!ledgerSheet) {
    ledgerSheet = ss.insertSheet('ledger');
    ledgerSheet.appendRow(['Timestamp', 'Asset', 'Action', 'Quantity', 'Price', 'Running P&L']);
  }

  // Create a trigger to update data every 2 hours if not already present
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
 * Main function to fetch latest prices and update the data sheet.
 * Also recalculates ledger P&L after updating prices.
 */
function updateData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('data');
  if (!sheet) {
    sheet = ss.insertSheet('data');
    sheet.appendRow(['Timestamp', 'BTC', 'ETH', 'SOL']);
  }

  // Clear previous data (but keep headers)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
  }

  try {
    var end = new Date();
    var start = new Date(end.getTime() - 24 * 60 * 60 * 1000); // 24 hours ago

    // Fetch candles for each asset
    var btc = getCandles('BTC-USD', start, end);
    var eth = getCandles('ETH-USD', start, end);
    var sol = getCandles('SOL-USD', start, end);

    // Build a map of timestamp -> prices
    var map = {};
    [btc, eth, sol].forEach(function(list, idx) {
      var name = idx === 0 ? 'BTC' : idx === 1 ? 'ETH' : 'SOL';
      list.forEach(function(rec) {
        var ts = rec[0].getTime();
        if (!map[ts]) {
          map[ts] = {timestamp: rec[0]};
        }
        map[ts][name] = rec[1];
      });
    });

    // Convert map to sorted rows
    var rows = Object.keys(map).sort().map(function(key) {
      var item = map[key];
      return [
        Utilities.formatDate(item.timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
        item.BTC || '',
        item.ETH || '',
        item.SOL || ''
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }
  } catch (e) {
    // Show error message in sheet
    sheet.getRange(2, 1).setValue('Error: ' + e.message);
  }

  // Format headers and columns
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.autoResizeColumns(1, 4);

  // Recalculate P&L with updated prices
  calculatePNL();
}

/**
 * Fetch candle data for a product between start and end dates.
 * Returns an array of [Date, closePrice].
 */
function getCandles(productId, start, end) {
  var url = 'https://api.exchange.coinbase.com/products/' + productId +
            '/candles?granularity=7200&start=' + start.toISOString() +
            '&end=' + end.toISOString();
  var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  if (response.getResponseCode() !== 200) {
    throw new Error('API request failed for ' + productId + ': ' + response.getContentText());
  }
  var data = JSON.parse(response.getContentText());
  data.sort(function(a, b) { return a[0] - b[0]; });
  return data.map(function(row) {
    return [new Date(row[0] * 1000), row[4]]; // timestamp and closing price
  });
}

/**
 * Calculate running P&L for the ledger based on latest prices.
 */
function calculatePNL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ledger = ss.getSheetByName('ledger');
  var dataSheet = ss.getSheetByName('data');
  if (!ledger || ledger.getLastRow() < 2 || !dataSheet || dataSheet.getLastRow() < 2) {
    return;
  }

  // Latest prices from the data sheet
  var lastRow = dataSheet.getLastRow();
  var pricesRow = dataSheet.getRange(lastRow, 2, 1, 3).getValues()[0];
  var prices = { BTC: pricesRow[0], ETH: pricesRow[1], SOL: pricesRow[2] };

  // Read ledger entries
  var entries = ledger.getRange(2, 1, ledger.getLastRow() - 1, 5).getValues();
  var results = [];
  var holdings = { BTC: 0, ETH: 0, SOL: 0 };
  var costs = { BTC: 0, ETH: 0, SOL: 0 };
  var realized = 0;
  var totalBuys = 0;

  entries.forEach(function(row) {
    var asset = row[1];
    var action = row[2];
    var qty = parseFloat(row[3]) || 0;
    var price = parseFloat(row[4]) || 0;

    if (action === 'Buy') {
      holdings[asset] += qty;
      costs[asset] += qty * price;
      totalBuys += qty * price;
    } else if (action === 'Sell') {
      var avgCost = holdings[asset] ? costs[asset] / holdings[asset] : 0;
      realized += qty * (price - avgCost);
      costs[asset] -= avgCost * qty;
      holdings[asset] -= qty;
    }

    var unrealized = 0;
    for (var a in holdings) {
      if (holdings[a]) {
        unrealized += holdings[a] * prices[a] - costs[a];
      }
    }
    results.push([realized + unrealized]);
  });

  if (results.length) {
    ledger.getRange(2, 6, results.length, 1).setValues(results);
  }

  var finalPnL = results.length ? results[results.length - 1][0] : 0;
  var returnPct = totalBuys ? (finalPnL / totalBuys) * 100 : 0;
  ledger.getRange(results.length + 3, 5, 1, 2).setValues([[
    'Total Return %', returnPct
  ]]);

  // Formatting
  ledger.getRange(1, 1, 1, 6).setFontWeight('bold');
  ledger.autoResizeColumns(1, 6);
}

/**
 * Triggered when the ledger sheet is edited to keep P&L updated.
 */
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === 'ledger') {
    calculatePNL();
  }
}
