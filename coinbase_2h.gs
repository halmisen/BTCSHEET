diff --git a/coinbase_2h.gs b/coinbase_2h.gs
index a7367a3264f11e2f9cfee854849d11de0521d942..720b93b3327f786bd773cee6bd3ef9976a13293a 100644
--- a/coinbase_2h.gs
+++ b/coinbase_2h.gs
@@ -1,43 +1,43 @@
 // Google Apps Script for fetching Coinbase 2h prices
 
 const GRANULARITY_SECONDS = 3600;   // 1h candle
 const TWO_HOUR_CANDLES   = 13;      // Data table keeps 13 two hour candles
 const DAILY_RESET_HOUR   = 8;       // daily sheet rollover time
 const HEADERS = ['Timestamp', 'BTC', 'ETH', 'SOL', 'Δ2h', 'Δ4h', 'Δ8h', 'Δ12h', 'Δ24h'];
 
 function formatTs(ts) {
   return Utilities.formatDate(new Date(ts * 1000), Session.getScriptTimeZone(),
     'yyyy-MM-dd HH:mm');
 }
 
 function getSheet() {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getSheetByName('Data');
   if (!sheet) {
     sheet = ss.insertSheet('Data');
-    sheet.appendRow(HEADERS);
+    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
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
diff --git a/coinbase_2h.gs b/coinbase_2h.gs
index a7367a3264f11e2f9cfee854849d11de0521d942..720b93b3327f786bd773cee6bd3ef9976a13293a 100644
--- a/coinbase_2h.gs
+++ b/coinbase_2h.gs
@@ -90,94 +90,100 @@ function fetchHistorical2hCandles(product, startIso, endIso) {
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
   var data = {};
   for (var i = 0; i < products.length; i++) {
     data[products[i]] = fetchRecent2hCandles(products[i], TWO_HOUR_CANDLES);
   }
 
+  var lastTs = '';
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
-    sheet.getRange(j + 2, 1, 1, 4).setValues([row]);
+    while (row.length < HEADERS.length) row.push('');
+    if (sheet.getLastColumn() < HEADERS.length) {
+      sheet.insertColumnsAfter(sheet.getLastColumn(), HEADERS.length - sheet.getLastColumn());
+    }
+    sheet.getRange(j + 2, 1, 1, HEADERS.length).setValues([row]);
+    lastTs = row[0];
   }
 
   var extra = sheet.getLastRow() - (TWO_HOUR_CANDLES + 1);
   if (extra > 0) {
     sheet.deleteRows(TWO_HOUR_CANDLES + 2, extra);
   }
-  return {timestamp: formatTs(Date.now()/1000)};
+  return { rows: TWO_HOUR_CANDLES, lastTs: lastTs };
 }
 
 function initHistory() {
   var sheet = getSheet();
   sheet.clear();
-  sheet.appendRow(HEADERS);
+  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
 
   var products = ['BTC-USD', 'ETH-USD', 'SOL-USD'];
   var data = {};
   for (var i = 0; i < products.length; i++) {
     data[products[i]] = fetchRecent2hCandles(products[i], TWO_HOUR_CANDLES);
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
-    for (var k = 0; k < 5; k++) row.push('');
-    sheet.appendRow(row);
+    while (row.length < HEADERS.length) row.push('');
+    sheet.getRange(j + 2, 1, 1, HEADERS.length).setValues([row]);
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
 
   var batchSeconds = 7200 * 350;
   var data = {};
   for (var i = 0; i < products.length; i++) {
     data[products[i]] = [];
   }
 
   for (var t = start.getTime(); t < end.getTime(); ) {
     var batchEnd = Math.min(t + batchSeconds * 1000, end.getTime());
     var sIso = new Date(t).toISOString();
diff --git a/coinbase_2h.gs b/coinbase_2h.gs
index a7367a3264f11e2f9cfee854849d11de0521d942..720b93b3327f786bd773cee6bd3ef9976a13293a 100644
--- a/coinbase_2h.gs
+++ b/coinbase_2h.gs
@@ -190,34 +196,39 @@ function backfillHistory(startDate, endDate) {
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
-  var dataSheet = getSheet();
-  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
-  var existing = ss.getSheetByName(today);
-  if (existing) ss.deleteSheet(existing);
-  dataSheet.copyTo(ss).setName(today);
-  dataSheet.clear();
-  dataSheet.appendRow(HEADERS);
+  var data = getSheet();
+  var today= Utilities.formatDate(new Date(), Session.getScriptTimeZone(),'yyyy-MM-dd');
+
+  // 归档旧表
+  var archived = data.copyTo(ss).setName(today);
+  SpreadsheetApp.flush();                      // 避免异步复制延迟
+
+  // 清空并重建 Data
+  data.clear();
+  data.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
+
+  // 立即填充 13 行
   update2hPrices();
 }
