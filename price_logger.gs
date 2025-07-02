/**
 * Crypto Price Logger for Google Sheets
 *
 * Provides utilities to fetch BTC, ETH and SOL prices from CoinGecko
 * and log them in a sheet named "Data". Run installTriggers() once
 * to create a two-hourly trigger that records prices automatically.
 */

// -----------------------------------------------------------------------------
// Configuration
// -----------------------------------------------------------------------------

/** Default sheet name used for logging */
const DATA_SHEET_NAME = 'Data';

/** Header row for the Data sheet */
const DATA_HEADERS = ['Timestamp', 'BTC', 'ETH', 'SOL'];

// -----------------------------------------------------------------------------
// Price fetching
// -----------------------------------------------------------------------------

/**
 * Fetch latest prices from CoinGecko.
 * @return {{BTC: number, ETH: number, SOL: number}}
 */
function fetchPrices() {
  const url = 'https://api.coingecko.com/api/v3/simple/price?ids=bitcoin,ethereum,solana&vs_currencies=usd';
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('CoinGecko request failed with HTTP ' + response.getResponseCode());
  }
  const data = JSON.parse(response.getContentText());
  return {
    BTC: data.bitcoin.usd,
    ETH: data.ethereum.usd,
    SOL: data.solana.usd
  };
}

// -----------------------------------------------------------------------------
// Sheet utilities
// -----------------------------------------------------------------------------

/**
 * Append prices with timestamp to the given sheet.
 * The sheet is created with headers if it does not exist.
 * @param {string} sheetName Name of the sheet to log into.
 * @param {{BTC:number,ETH:number,SOL:number}} prices Price map from fetchPrices().
 */
function logPrices(sheetName, prices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, DATA_HEADERS.length)
      .setValues([DATA_HEADERS])
      .setFontWeight('bold');
  }

  const headers = sheet.getRange(1, 1, 1, DATA_HEADERS.length).getValues()[0];
  if (headers.join() !== DATA_HEADERS.join()) {
    sheet.clear();
    sheet.getRange(1, 1, 1, DATA_HEADERS.length)
      .setValues([DATA_HEADERS])
      .setFontWeight('bold');
  }

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([ts, prices.BTC, prices.ETH, prices.SOL]);
}

// -----------------------------------------------------------------------------
// Main entry and trigger management
// -----------------------------------------------------------------------------

/**
 * Fetch prices and log them. Sends an email on failure.
 */
function fetchAndLogPrices() {
  try {
    const prices = fetchPrices();
    logPrices(DATA_SHEET_NAME, prices);
  } catch (err) {
    const user = Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(user, 'Crypto price logger error', String(err));
    throw err;
  }
}

/**
 * Remove existing time triggers and create a new 2h trigger for fetchAndLogPrices.
 */
function installTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('fetchAndLogPrices')
    .timeBased()
    .everyHours(2)
    .create();
}

