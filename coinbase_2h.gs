// Google Apps Script for fetching Coinbase 2h prices

const GRANULARITY_SECONDS = 3600;   // 1h candle
const TWO_HOUR_CANDLES   = 13;      // Data table keeps 13 two hour candles
const DAILY_RESET_HOUR   = 8;       // daily sheet rollover time
const HEADERS = ['Timestamp', 'BTC', 'ETH', 'SOL', 'Δ2h', 'Δ4h', 'Δ8h', 'Δ12h', 'Δ24h'];
// Additional columns for latest change summary
const CHANGE_HEADERS = [
  'BTC Δ2h', 'BTC Δ4h', 'BTC Δ12h', 'BTC Δ24h',
  'ETH Δ2h', 'ETH Δ4h', 'ETH Δ12h', 'ETH Δ24h',
  'SOL Δ2h', 'SOL Δ4h', 'SOL Δ12h', 'SOL Δ24h'
];

function formatTs(ts) {
  return Utilities.formatDate(new Date(ts * 1000), Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm');
}

function formatChange(curr, prev) {
  if (curr == null || curr === '' || prev == null || prev === '') return '';
  var diff = curr - prev;
  var pct = prev ? (diff / prev) * 100 : 0;
  return diff.toFixed(1) + ' (' + pct.toFixed(2) + '%)';
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  if (!sheet) {
    sheet = ss.insertSheet('Data');
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function fetchLatestCandle(product) {
  var url = 'https://api.exchange.coinbase.com/products/' + product + '/candles?granularity=' + GRANULARITY_SECONDS + '&limit=2';
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

function fetchLatest2hCandles(product, limit){
  const url = `https://api.exchange.coinbase.com/products/${product}/candles`+
              `?granularity=3600&limit=${limit*2}`;
  const res = UrlFetchApp.fetch(url,{headers:{Accept:'application/json'}});
  if(res.getResponseCode()!=200) throw res.getContentText();
  const arr = JSON.parse(res).sort((a,b)=>b[0]-a[0]); // 最新在前
  const out = [];
  for(let i=0;i+1< arr.length && out.length<limit;i+=2){
     const c1 = arr[i], c2 = arr[i+1];
     out.push([             // [ts, low, high, open, close]
       c2[0],               // 取较早蜡烛的时间戳
       Math.min(c1[1],c2[1]),
       Math.max(c1[2],c2[2]),
       c2[3],
       c1[4]
     ]);
  }
  return out.reverse();     // 升序写表
}

function fetchHistorical2hCandles(product, startIso, endIso) {
  var url = 'https://api.exchange.coinbase.com/products/' + product +
    '/candles?granularity=3600&start=' + startIso + '&end=' + endIso;
  var options = {headers: {Accept: 'application/json'}};
  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('HTTP ' + response.getResponseCode() + ' for history ' + product);
      return [];
    }
    var arr = JSON.parse(response.getContentText());
    if (!arr) return [];
    arr.sort(function(a, b) { return a[0] - b[0]; });
    var out = [];
    for (var i = 0; i + 1 < arr.length; i += 2) {
      var c1 = arr[i], c2 = arr[i + 1];
      out.push([
        c1[0],
        Math.min(c1[1], c2[1]),
        Math.max(c1[2], c2[2]),
        c1[3],
        c2[4]
      ]);
    }
    return out;
  } catch (e) {
    Logger.log('Error fetching history for ' + product + ': ' + e);
    return [];
  }
}

function update2hPrices() {
  var sheet = getSheet();

  var products = ['BTC-USD', 'ETH-USD', 'SOL-USD'];
  var data = {};
  for (var i = 0; i < products.length; i++) {
    data[products[i]] = fetchLatest2hCandles(products[i], TWO_HOUR_CANDLES);
  }

  var lastTs = '';
  for (var j = 0; j < TWO_HOUR_CANDLES; j++) {
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
    var btc = data['BTC-USD'];
    var deltas = [];
    var offs = [1, 2, 4, 6, 12];
    for (var d = 0; d < offs.length; d++) {
      var o = offs[d];
      if (j >= o && btc[j] && btc[j - o] && btc[j][4] != null && btc[j - o][4] != null) {
        deltas.push(parseFloat(btc[j][4]) - parseFloat(btc[j - o][4]));
      } else {
        deltas.push('');
      }
    }
    row = row.concat(deltas);
    while (row.length < HEADERS.length) row.push('');
    if (sheet.getLastColumn() < HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), HEADERS.length - sheet.getLastColumn());
    }
    sheet.getRange(j + 2, 1, 1, HEADERS.length).setValues([row]);
    lastTs = row[0];
  }

  var extra = sheet.getLastRow() - (TWO_HOUR_CANDLES + 1);
  if (extra > 0) {
    sheet.deleteRows(TWO_HOUR_CANDLES + 2, extra);
  }
  // update summary row with newest values
  updateLatestChanges();
  return { rows: TWO_HOUR_CANDLES, lastTs: lastTs };
}

function initHistory() {
  var sheet = getSheet();
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  var products = ['BTC-USD', 'ETH-USD', 'SOL-USD'];
  var data = {};
  for (var i = 0; i < products.length; i++) {
    data[products[i]] = fetchLatest2hCandles(products[i], TWO_HOUR_CANDLES);
  }
  for (var j = 0; j < TWO_HOUR_CANDLES; j++) {
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
    var btc = data['BTC-USD'];
    var deltas = [];
    var offs = [1, 2, 4, 6, 12];
    for (var d = 0; d < offs.length; d++) {
      var o = offs[d];
      if (j >= o && btc[j] && btc[j - o] && btc[j][4] != null && btc[j - o][4] != null) {
        deltas.push(parseFloat(btc[j][4]) - parseFloat(btc[j - o][4]));
      } else {
        deltas.push('');
      }
    }
    row = row.concat(deltas);
    while (row.length < HEADERS.length) row.push('');
    sheet.getRange(j + 2, 1, 1, HEADERS.length).setValues([row]);
  }
}

function backfillHistory(startDate, endDate) {
  var products = ['BTC-USD', 'ETH-USD', 'SOL-USD'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'History_' + startDate + '_to_' + endDate;
  var exist = ss.getSheetByName(sheetName);
  if (exist) ss.deleteSheet(exist);
  var sheet = ss.insertSheet(sheetName);
  sheet.appendRow(HEADERS);

  var start = new Date(startDate + 'T00:00:00Z');
  var end = new Date(endDate + 'T00:00:00Z');
  end.setDate(end.getDate() + 1);

  var batchSeconds = 3600 * 700; // fetch 350 two-hour periods as 700 one-hour candles
  var data = {};
  for (var i = 0; i < products.length; i++) {
    data[products[i]] = [];
  }

  for (var t = start.getTime(); t < end.getTime(); ) {
    var batchEnd = Math.min(t + batchSeconds * 1000, end.getTime());
    var sIso = new Date(t).toISOString();
    var eIso = new Date(batchEnd).toISOString();
    for (var p = 0; p < products.length; p++) {
      data[products[p]] = data[products[p]].concat(
        fetchHistorical2hCandles(products[p], sIso, eIso)
      );
      Utilities.sleep(350);
    }
    t = batchEnd;
  }

  var allTs = {};
  products.forEach(function(pr) {
    data[pr].forEach(function(c) {
      if (!allTs[c[0]]) allTs[c[0]] = {};
      allTs[c[0]][pr] = parseFloat(c[4]);
    });
  });

  var tsList = Object.keys(allTs).map(Number).sort(function(a, b) { return a - b; });
  tsList.forEach(function(ts) {
    var row = [formatTs(ts)];
    products.forEach(function(pr) {
      row.push(allTs[ts][pr] != null ? allTs[ts][pr] : 'N/A');
    });
    for (var k = 0; k < 5; k++) row.push('');
    sheet.appendRow(row);
  });
}

function rolloverDailySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = getSheet();
  var today= Utilities.formatDate(new Date(), Session.getScriptTimeZone(),'yyyy-MM-dd');

  // 归档旧表
  var archived = data.copyTo(ss).setName(today);
  SpreadsheetApp.flush();                      // 避免异步复制延迟

  // 清空并重建 Data
  data.clear();
  data.appendRow(HEADERS);

  // 立即填充 13 行
  update2hPrices();
}

function updateLatestChanges() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // nothing to summarise

  var lastLabel = sheet.getRange(lastRow, 1).getValue();
  var dataRow = lastLabel === 'Summary' ? lastRow - 1 : lastRow;

  // ensure summary headers exist and there are enough columns
  var startCol = HEADERS.length + 1;
  var required = startCol + CHANGE_HEADERS.length - 1;
  if (sheet.getLastColumn() < required) {
    sheet.insertColumnsAfter(sheet.getLastColumn(), required - sheet.getLastColumn());
  }
  sheet.getRange(1, startCol, 1, CHANGE_HEADERS.length).setValues([CHANGE_HEADERS]);

  // columns holding BTC, ETH, SOL - filter in case sheet is missing some
  var cols = [2, 3, 4].filter(function(c) { return c <= sheet.getLastColumn(); });
  var offsets = [1, 2, 6, 12]; // 2h, 4h, 12h, 24h
  var out = [];

  for (var c = 0; c < cols.length; c++) {
    var curr = sheet.getRange(dataRow, cols[c]).getValue();
    for (var i = 0; i < offsets.length; i++) {
      var prevRow = dataRow - offsets[i];
      if (prevRow >= 2) {
        var prev = sheet.getRange(prevRow, cols[c]).getValue();
        out.push(formatChange(curr, prev));
      } else {
        out.push('');
      }
    }
  }

  var summaryRow = dataRow + 1;
  if (lastLabel !== 'Summary') {
    sheet.insertRowsAfter(dataRow, 1);
  }
  sheet.getRange(summaryRow, 1, 1, required).clearContent();
  sheet.getRange(summaryRow, 1).setValue('Summary');
  sheet.getRange(summaryRow, startCol, 1, CHANGE_HEADERS.length).setValues([out]);
}
function updateChangeColumns() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // nothing to do

  var offsets = [1, 2, 6, 12]; // number of rows back for 2h,4h,12h,24h
  var suffixes = ['_Δ2h', '_Δ4h', '_Δ12h', '_Δ24h'];

  var col = 2; // first coin column after Timestamp
  while (col <= sheet.getLastColumn()) {
    var header = String(sheet.getRange(1, col).getValue());
    if (!header || header.indexOf('Δ') >= 0) {
      col++;
      continue;
    }
    // ensure four delta columns exist after the coin column
    for (var i = 0; i < suffixes.length; i++) {
      var expectCol = col + i + 1;
      if (expectCol > sheet.getLastColumn()) {
        sheet.insertColumnAfter(sheet.getLastColumn());
      }
      var h = sheet.getRange(1, expectCol).getValue();
      if (h !== header + suffixes[i]) {
        sheet.insertColumnAfter(expectCol - 1);
        sheet.getRange(1, expectCol).setValue(header + suffixes[i]);
      }
    }
    // recalc after possible insertions
    var last = sheet.getRange(lastRow, col).getValue();
    for (var i = 0; i < offsets.length; i++) {
      var targetCol = col + i + 1;
      sheet.getRange(2, targetCol, lastRow - 1).clearContent();
      var prevRow = lastRow - offsets[i];
      var val = '';
      if (prevRow >= 2) {
        var prev = sheet.getRange(prevRow, col).getValue();
        val = formatChange(last, prev);
      }
      if (val) {
        sheet.getRange(lastRow, targetCol).setValue(val);
      }
    }
    col += suffixes.length + 1; // move past coin and its delta cols
  }
}
