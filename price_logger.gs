/**
 * Crypto Price Logger for Google Sheets
 *
 * Provides utilities to fetch BTC, ETH and SOL prices from Coinbase
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
 * Fetch latest spot prices from Coinbase with retries.
 * @return {{BTC: number, ETH: number, SOL: number}}
 */
function fetchPrices() {
  const ids = ['BTC', 'ETH', 'SOL'];
  const prices = {};

  ids.forEach(id => {
    let attempt = 0;
    while (true) {
      attempt++;
      try {
        // Coinbase spot price endpoint, e.g. https://api.coinbase.com/v2/prices/BTC-USD/spot
        const url = `https://api.coinbase.com/v2/prices/${id}-USD/spot`;
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (response.getResponseCode() !== 200) {
          throw new Error('HTTP ' + response.getResponseCode());
        }

        const json = JSON.parse(response.getContentText());
        prices[id] = parseFloat(json.data.amount);
        break; // success
      } catch (err) {
        if (attempt >= 3) {
          // After 3 attempts give up and rethrow
          throw new Error('Failed to fetch ' + id + ': ' + err);
        }
        // Exponential backoff before retrying
        Utilities.sleep(Math.pow(2, attempt - 1) * 1000);
      }
    }
  });

  return prices;
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

  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Write headers on first use
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, DATA_HEADERS.length)
      .setValues([DATA_HEADERS])
      .setFontWeight('bold');
  }

  // Append timestamp and prices to next row
  sheet.appendRow([new Date(), prices.BTC, prices.ETH, prices.SOL]);
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
    const subject = 'Crypto price logger error';
    // Notify the current user about the failure
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, String(err));
    throw err;
  }
}

/**
 * Remove existing time triggers and create a new 2h trigger for fetchAndLogPrices.
 */
function installTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  // Delete all existing triggers to avoid duplicates
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  // Create a new time-based trigger to run every 2 hours
  ScriptApp.newTrigger('fetchAndLogPrices')
    .timeBased()
    .everyHours(2)
    .create();
}

// -----------------------------------------------------------------------------
// Deployment notes
// -----------------------------------------------------------------------------

// After deploying the project, run installTriggers() once manually to enable
// automatic price logging every two hours.


