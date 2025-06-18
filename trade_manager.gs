// Google Apps Script utilities for a paper trading ledger

/** Headers used in the Ledger sheet */
const LEDGER_HEADERS = [
  'Trade ID',
  'Trade Time',
  'Symbol',
  'Side',
  'Price',
  'Quantity',
  'Trade Amount',
  'Running Position',
  'Average Cost',
  'Floating P&L'
];

/** Return the Ledger sheet, creating it if necessary */
function getLedgerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Ledger');
  if (!sheet) {
    sheet = ss.insertSheet('Ledger');
    sheet.appendRow(LEDGER_HEADERS);
    setupLedgerValidation();
  }
  return sheet;
}

/** Setup dropdowns for Symbol and Side columns */
function setupLedgerValidation() {
  var sheet = getLedgerSheet();
  var last = sheet.getMaxRows() - 1;
  var symbolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['BTC','ETH','SOL'], true)
    .build();
  sheet.getRange(2, 3, last).setDataValidation(symbolRule);

  var sideRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Buy','Sell'], true)
    .build();
  sheet.getRange(2, 4, last).setDataValidation(sideRule);
}

/** Find price for symbol at the given timestamp in Data sheet */
function findPrice(symbol, timeStr) {
  var sheet = getSheet(); // Data sheet from coinbase_2h.gs
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  var timeVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var col = {BTC:2, ETH:3, SOL:4}[symbol];
  if (!col) return null;
  for (var i = 0; i < timeVals.length; i++) {
    if (timeVals[i][0] == timeStr) {
      return sheet.getRange(i + 2, col).getValue();
    }
  }
  return null;
}

/** Return the latest price for each symbol from the Data sheet */
function getLatestPrices() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  var prices = {BTC:0, ETH:0, SOL:0};
  if (lastRow <= 1) return prices;
  prices.BTC = sheet.getRange(lastRow, 2).getValue();
  prices.ETH = sheet.getRange(lastRow, 3).getValue();
  prices.SOL = sheet.getRange(lastRow, 4).getValue();
  return prices;
}

/** Recalculate ledger formulas for all rows and update summary */
function recomputeLedger() {
  var sheet = getLedgerSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var col = {id:0, time:1, sym:2, side:3, price:4, qty:5,
             amt:6, pos:7, avg:8, pnl:9};

  var pos = {BTC:0, ETH:0, SOL:0};
  var avg = {BTC:0, ETH:0, SOL:0};
  var lastPrices = getLatestPrices();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var sym = row[col.sym];
    if (!sym) continue;
    var price = parseFloat(row[col.price]);
    var qty = parseFloat(row[col.qty]);
    if (isNaN(price) || isNaN(qty)) continue;
    var sign = row[col.side] == 'Buy' ? 1 : -1;
    var signedQty = qty * sign;

    var prevPos = pos[sym];
    var prevAvg = avg[sym];
    var newPos = prevPos + signedQty;
    var newAvg = prevAvg;
    if (sign > 0) { // buy
      newAvg = (prevAvg * Math.abs(prevPos) + price * qty) / Math.abs(newPos);
    } else { // sell
      if (Math.sign(prevPos) == Math.sign(newPos) && prevPos != 0) {
        newAvg = prevAvg;
      } else if (newPos == 0) {
        newAvg = 0;
      } else {
        newAvg = price;
      }
    }

    pos[sym] = newPos;
    avg[sym] = newAvg;

    row[col.amt] = price * qty * sign;
    row[col.pos] = newPos;
    row[col.avg] = newAvg;
    row[col.pnl] = (lastPrices[sym] - newAvg) * newPos;

    sheet.getRange(i + 1, col.amt + 1, 1, 4)
      .setValues([[row[col.amt], row[col.pos], row[col.avg], row[col.pnl]]]);
  }

  // Write summary block
  var start = sheet.getLastRow() + 2;
  var out = [['Symbol','Position','Avg Cost','Floating P&L']];
  ['BTC','ETH','SOL'].forEach(function(s) {
    out.push([s, pos[s], avg[s], (lastPrices[s] - avg[s]) * pos[s]]);
  });
  sheet.getRange(start, 1, out.length, out[0].length).clearContent();
  sheet.getRange(start, 1, out.length, out[0].length).setValues(out);
}

/** Handle edits on the Ledger sheet */
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() != 'Ledger') return;
  var row = e.range.getRow();
  if (row <= 1) return;

  var values = sheet.getRange(row, 1, 1, LEDGER_HEADERS.length).getValues()[0];
  var id = values[0];
  var time = values[1];
  var sym = values[2];
  var side = values[3];
  var price = values[4];
  var qty = values[5];

  if (!id && (time && sym)) {
    var nextId = sheet.getRange(sheet.getLastRow(), 1).getValue();
    nextId = nextId ? nextId + 1 : 1;
    sheet.getRange(row, 1).setValue(nextId);
  }

  if (!price && time && sym) {
    var p = findPrice(sym, time);
    if (p != null) sheet.getRange(row, 5).setValue(p);
  }

  if (time && sym && side && qty) {
    recomputeLedger();
  }
}

/** Convenience to create the Ledger sheet and validation rules */
function initLedger() {
  var sheet = getLedgerSheet();
  sheet.clear();
  sheet.appendRow(LEDGER_HEADERS);
  setupLedgerValidation();
}
// Google Apps Script for fetching Coinbase 2h prices

const GRANULARITY_SECONDS = 3600;   // 1h candle
const TWO_HOUR_CANDLES   = 13;      // Data table keeps 13 two hour candles
const DAILY_RESET_HOUR   = 8;       // daily sheet rollover time
// Columns holding raw prices only. Change columns are generated dynamically.
const HEADERS = ['Timestamp', 'BTC', 'ETH', 'SOL'];

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
  // update summary columns with newest values
  refreshLatestChanges();
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
    if (sheet.getLastColumn() < HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), HEADERS.length - sheet.getLastColumn());
    }
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

/**
 * Rebuild the price change summary table below the main data.
 * The table shows the latest price difference and percent change compared to
 * the rows 2h, 4h, 12h and 24h earlier. Existing summary rows are cleared
 * before the new table is written.
 */
function refreshLatestChanges() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // nothing to summarise

  // Detect coin columns after Timestamp
  var headerValues = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1).getValues()[0];
  var coins = [];
  for (var i = 0; i < headerValues.length; i++) {
    if (headerValues[i]) coins.push({ name: headerValues[i], col: i + 2 });
  }

  var offsets = [1, 2, 6, 12];
  var suffixes = ['Δ2h', 'Δ4h', 'Δ12h', 'Δ24h'];

  var startRow = lastRow + 2; // one blank row after data
  var rows = coins.length + 1; // header + coins
  var cols = suffixes.length + 1;
  sheet.getRange(startRow, 1, rows, cols).clearContent();

  var values = [];
  values.push([''].concat(suffixes));

  coins.forEach(function(c) {
    var row = [c.name];
    var current = sheet.getRange(lastRow, c.col).getValue();
    offsets.forEach(function(off) {
      var prevRow = lastRow - off;
      if (prevRow >= 2) {
        var prev = sheet.getRange(prevRow, c.col).getValue();
        row.push(formatChange(current, prev));
      } else {
        row.push('');
      }
    });
    values.push(row);
  });

  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
  sheet.getRange(1, 1, 1, values[0].length)
       .copyFormatToRange(sheet, 1, values[0].length, startRow, startRow);
}
