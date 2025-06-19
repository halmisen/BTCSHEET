/**
 * Simple cryptocurrency price logger for Google Sheets.
 *
 * Adds a custom menu "Crypto Tools" with an option to fetch the
 * latest BTC, ETH and SOL spot prices in USD from Coinbase. The
 * prices are appended to a sheet named "Data" along with the
 * current timestamp each time the menu item is clicked.
 *
 * The script automatically creates the "Data" sheet and its
 * header row if they do not already exist.
 */

// -----------------------------------------------------------------------------
// Configuration
// -----------------------------------------------------------------------------

/** Name of the sheet used to store price data */
const DATA_SHEET_NAME = 'Data';

/** Header row for the Data sheet */
const DATA_HEADERS = ['Timestamp', 'BTC', 'ETH', 'SOL'];

// -----------------------------------------------------------------------------
// Menu setup
// -----------------------------------------------------------------------------

/**
 * Creates the custom menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Crypto Tools')
    .addItem('Fetch Latest Prices', 'fetchLatestPrices')
    .addToUi();
  ensureDataSheet();
}

// -----------------------------------------------------------------------------
// Sheet utilities
// -----------------------------------------------------------------------------

/**
 * Ensures the Data sheet exists with the correct header row.
 * If the sheet is missing it will be created. If the header row
 * does not match DATA_HEADERS it will be replaced.
 */
function ensureDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET_NAME);
  }

  // Check if the first row matches the expected headers
  const firstRow = sheet.getRange(1, 1, 1, DATA_HEADERS.length).getValues()[0];
  const headersMatch = firstRow.join() === DATA_HEADERS.join();

  if (!headersMatch) {
    sheet.clear();
    sheet.getRange(1, 1, 1, DATA_HEADERS.length)
         .setValues([DATA_HEADERS])
         .setFontWeight('bold');
  }
}

// -----------------------------------------------------------------------------
// Price fetching utilities
// -----------------------------------------------------------------------------

/**
 * Fetches the current USD spot price for the given symbol from Coinbase.
 * Returns a number on success or the string 'ERROR' if the request fails.
 */
function fetchPrice(symbol) {
  const url = `https://api.coinbase.com/v2/prices/${symbol}-USD/spot`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const price = parseFloat(json.data.amount);
      if (!isNaN(price)) {
        return price;
      }
    }
    Logger.log(`Error fetching ${symbol}: HTTP ${response.getResponseCode()}`);
  } catch (err) {
    Logger.log(`Fetch failed for ${symbol}: ${err}`);
  }
  return 'ERROR';
}

// -----------------------------------------------------------------------------
// Main price logging function
// -----------------------------------------------------------------------------

/**
 * Fetches latest BTC, ETH and SOL prices and appends them to the Data sheet.
 */
function fetchLatestPrices() {
  ensureDataSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const btc = fetchPrice('BTC');
  const eth = fetchPrice('ETH');
  const sol = fetchPrice('SOL');

  sheet.appendRow([timestamp, btc, eth, sol]);
}

// -----------------------------------------------------------------------------
// End of script
// -----------------------------------------------------------------------------

