// Google Apps Script for fetching Coinbase 2h prices

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
  var url = 'https://api.exchange.coinbase.com/products/' + product + '/candles?granularity=7200&limit=1';
  var options = {headers: {Accept: 'application/json'}};
  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('HTTP ' + response.getResponseCode() + ' for ' + product);
      return null;
    }
    var data = JSON.parse(response.getContentText());
    if (!data || data.length === 0) {
      Logger.log('Empty data for ' + product);
      return null;
    }
    return data[0];
  } catch (e) {
    Logger.log('Error fetching ' + product + ': ' + e);
    return null;
  }
}

function fetchCandles(product, limit) {
  var url = 'https://api.exchange.coinbase.com/products/' + product + '/candles?granularity=7200&limit=' + limit;
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
    return data;
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
      timestamp = candle[0] * 1000;
    }
    prices.push(candle ? parseFloat(candle[4]) : 'N/A');
  }
  if (!timestamp) {
    return {timestamp: null, prices: prices};
  }
  var lastRow = sheet.getLastRow();
  var lastTimestamp = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 0;
  if (timestamp > lastTimestamp) {
    sheet.appendRow([timestamp].concat(prices));
  }
  return {timestamp: timestamp, prices: prices};
}

function initHistory(limit) {
  limit = limit || 100;
  var sheet = getSheet();
  sheet.clear();
  sheet.appendRow(['Timestamp', 'BTC', 'ETH', 'SOL']);

  var products = ['BTC-USD', 'ETH-USD', 'SOL-USD'];
  var data = {};
  for (var i = 0; i < products.length; i++) {
    data[products[i]] = fetchCandles(products[i], limit);
  }
  for (var j = 0; j < limit; j++) {
    var ts = data['BTC-USD'][j] ? data['BTC-USD'][j][0] * 1000 : '';
    var row = [
      ts,
      data['BTC-USD'][j] ? parseFloat(data['BTC-USD'][j][4]) : 'N/A',
      data['ETH-USD'][j] ? parseFloat(data['ETH-USD'][j][4]) : 'N/A',
      data['SOL-USD'][j] ? parseFloat(data['SOL-USD'][j][4]) : 'N/A'
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
