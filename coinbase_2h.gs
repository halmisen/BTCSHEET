// Google Apps Script for fetching Coinbase 2h prices

var HISTORY_ROWS = 12; // number of 2h candles to fetch when initializing

function formatTs(ts) {
  return Utilities.formatDate(new Date(ts * 1000), Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm');
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  if (!sheet) {
    sheet = ss.insertSheet('Data');
    sheet.appendRow(['Timestamp', 'BTC', 'ETH', 'SOL']);
  }
  return sheet;
}

function fetchLatestCandle(product) {
  var url = 'https://api.exchange.coinbase.com/products/' + product + '/candles?granularity=3600&limit=2';
  var options = {headers: {Accept: 'application/json'}};
  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('HTTP ' + response.getResponseCode() + ' for ' + product);
      return null;
    }
    var data = JSON.parse(response.getContentText());
    if (!data || data.length < 2) {
      Logger.log('Insufficient data for ' + product);
      return null;
    }
    data.sort(function(a, b) { return a[0] - b[0]; });
    var c1 = data[data.length - 2];
    var c2 = data[data.length - 1];
    return [
      c1[0],
      Math.min(c1[1], c2[1]),
      Math.max(c1[2], c2[2]),
      c1[3],
      c2[4],
      (c1[5] || 0) + (c2[5] || 0)
    ];
  } catch (e) {
    Logger.log('Error fetching ' + product + ': ' + e);
    return null;
  }
}

function fetchCandles(product, limit) {
  var url = 'https://api.exchange.coinbase.com/products/' + product + '/candles?granularity=3600&limit=' + (limit * 2);
  var options = {headers: {Accept: 'application/json'}};
  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('HTTP ' + response.getResponseCode() + ' for history ' + product);
      return [];
    }
    var data = JSON.parse(response.getContentText());
    if (!data) return [];
    data.sort(function(a, b) { return a[0] - b[0]; });
    var result = [];
    for (var i = 0; i + 1 < data.length && result.length < limit; i += 2) {
      var c1 = data[i];
      var c2 = data[i + 1];
      result.push([
        c1[0],
        Math.min(c1[1], c2[1]),
        Math.max(c1[2], c2[2]),
        c1[3],
        c2[4],
        (c1[5] || 0) + (c2[5] || 0)
      ]);
    }
    return result;
  } catch (e) {
    Logger.log('Error fetching history for ' + product + ': ' + e);
    return [];
  }
}

function update2hPrices() {
  var sheet = getSheet();
  var products = ['BTC-USD', 'ETH-USD', 'SOL-USD'];
  var prices = [];
  var timestamp = null;
  for (var i = 0; i < products.length; i++) {
    var candle = fetchLatestCandle(products[i]);
    if (!timestamp && candle) {
      timestamp = candle[0];
    }
    if (candle && candle[4] != null) {
      prices.push(parseFloat(candle[4]));
    } else {
      Logger.log('Missing price for ' + products[i]);
      prices.push('N/A');
    }
  }
  if (!timestamp) {
    return {timestamp: null, prices: prices};
  }
  var formatted = formatTs(timestamp);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    initHistory();
    lastRow = sheet.getLastRow();
  }
  var lastTimestampStr = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : '';
  var lastTimestampMs = lastTimestampStr ? new Date(lastTimestampStr).getTime() : 0;
  if (timestamp * 1000 > lastTimestampMs) {
    sheet.appendRow([formatted].concat(prices));
  }
  return {timestamp: formatted, prices: prices};
}

function initHistory() {
  var sheet = getSheet();
  sheet.clear();
  sheet.appendRow(['Timestamp', 'BTC', 'ETH', 'SOL']);

  var products = ['BTC-USD', 'ETH-USD', 'SOL-USD'];
  var data = {};
  for (var i = 0; i < products.length; i++) {
    data[products[i]] = fetchCandles(products[i], HISTORY_ROWS);
  }
  for (var j = 0; j < HISTORY_ROWS; j++) {
    var ts = data['BTC-USD'][j] ? formatTs(data['BTC-USD'][j][0]) : '';
    var row = [
      ts,
      data['BTC-USD'][j] && data['BTC-USD'][j][4] != null ?
        parseFloat(data['BTC-USD'][j][4]) : 'N/A',
      data['ETH-USD'][j] && data['ETH-USD'][j][4] != null ?
        parseFloat(data['ETH-USD'][j][4]) : 'N/A',
      data['SOL-USD'][j] && data['SOL-USD'][j][4] != null ?
        parseFloat(data['SOL-USD'][j][4]) : 'N/A'
    ];
    sheet.appendRow(row);
  }
}

(function() {
  try {
    var result = update2hPrices();
    console.log(result);
  } catch (e) {
    console.log(e);
  }
})();
