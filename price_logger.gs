/**
 * Crypto Price Logger for Google Sheets
 *
 * Provides utilities to fetch BTC, ETH and SOL spot prices from the
 * Coinbase public API and log them in a sheet named "Data". Run
 * installTriggers() once to create a two-hourly trigger that records
 * prices automatically.
 */

// -----------------------------------------------------------------------------
// Configuration
// -----------------------------------------------------------------------------

/** Default sheet name used for logging */
const DATA_SHEET_NAME = 'Data';

/** Header row for the Data sheet */
const DATA_HEADERS = ['Timestamp', 'BTC', 'ETH', 'SOL'];

/** List of crypto symbols to retrieve from Coinbase */
const CRYPTO_IDS = ['BTC', 'ETH', 'SOL'];

// -----------------------------------------------------------------------------
// Price fetching
// -----------------------------------------------------------------------------

/**
 * Fetch latest spot prices from Coinbase.
 * Retries up to 3 times with exponential backoff on failures.
 * @return {{BTC: number, ETH: number, SOL: number}}
 */
function fetchPrices() {
  const prices = {};
  const base = 'https://api.coinbase.com/v2/prices/';
  const maxAttempts = 3;

  CRYPTO_IDS.forEach(id => {
    let attempt = 0;
    while (attempt < maxAttempts) {
      try {
        const url = base + id + '-USD/spot';
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) {
          const json = JSON.parse(resp.getContentText());
          prices[id] = parseFloat(json.data.amount);
          break; // success
        }
        throw new Error('HTTP ' + resp.getResponseCode());
      } catch (err) {
        attempt++;
        if (attempt >= maxAttempts) {
          throw new Error('Failed to fetch ' + id + ' price: ' + err);
        }
        // Exponential backoff: 1s, 2s, 4s
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

  // Create the sheet if it does not exist
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Initialize header row when sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(DATA_HEADERS);
  }

  // Append the latest timestamp and price values
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
    MailApp.sendEmail(
      Session.getEffectiveUser(),
      'Crypto price logger error',
      String(err)
    );
    throw err;
  }
}

/**
 * Remove existing time triggers and create a new 2h trigger for fetchAndLogPrices.
 */
function installTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  // Remove all existing time-driven triggers
  triggers.forEach(t => {
    if (t.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new trigger that runs every 2 hours
  ScriptApp.newTrigger('fetchAndLogPrices')
    .timeBased()
    .everyHours(2)
    .create();
}

// After deploying, run installTriggers() manually once to start
// the automated 2h price logging.

