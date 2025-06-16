// Google Apps Script for fetching Binance 2h prices

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  if (!sheet) {
    sheet = ss.insertSheet('Data');
    sheet.appendRow(['Timestamp', 'BTC', 'ETH', 'SOL']);
  }
  return sheet;
}

function fetchLatestKline(symbol) {
  var url = 'https://api.binance.com/api/v3/klines?symbol=' + symbol + '&interval=2h&limit=1';
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  return data[0];
}

function update2hPrices() {
  var sheet = getSheet();
  var symbols = ['BTCUSDT', 'ETHUSDT', 'SOLUSDT'];
  var prices = [];
  var timestamp = null;
  for (var i = 0; i < symbols.length; i++) {
    var kline = fetchLatestKline(symbols[i]);
    if (!timestamp) {
      timestamp = kline[0];
    }
    prices.push(parseFloat(kline[4]));
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

  var symbols = ['BTCUSDT', 'ETHUSDT', 'SOLUSDT'];
  var data = {};
  for (var i = 0; i < symbols.length; i++) {
    var url = 'https://api.binance.com/api/v3/klines?symbol=' + symbols[i] + '&interval=2h&limit=' + limit;
    var response = UrlFetchApp.fetch(url);
    data[symbols[i]] = JSON.parse(response.getContentText());
  }

  for (var j = 0; j < limit; j++) {
    var ts = data['BTCUSDT'][j][0];
    var row = [
      ts,
      parseFloat(data['BTCUSDT'][j][4]),
      parseFloat(data['ETHUSDT'][j][4]),
      parseFloat(data['SOLUSDT'][j][4])
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
