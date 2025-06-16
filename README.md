# BTCSHEET

This repository contains sample Google Apps Script code to fetch cryptocurrency prices from Coinbase every 2 hours.

The `coinbase_2h.gs` script exposes several utilities:

- `update2hPrices()` rebuilds the `Data` sheet with the latest **13** two hour candles for BTC-USD, ETH-USD and SOL-USD.
- `rolloverDailySheet()` copies the current `Data` sheet to a new sheet named by date and then refreshes `Data` for the new day.
- `backfillHistory(start, end)` downloads historical two hour candles between two dates and stores them in a sheet named `History_<start>_to_<end>`.

The data is retrieved using the public Coinbase API endpoint:

```
https://api.exchange.coinbase.com/products/BTC-USD/candles?granularity=7200&limit=1
```

See the official documentation for details: <https://docs.cloud.coinbase.com/exchange/reference/exchangerestapi_getproductcandles>.

The public API allows about 3 requests per second, so this script only queries a few products at low frequency.

After deploying the script you can run `update2hPrices()` manually once. If the
sheet is still empty this will call `initHistory()` and backfill the latest
history before appending the newest prices.

```js
// 一次性回填 2025-01-01 ~ 2025-03-31
backfillHistory('2025-01-01','2025-03-31');
```

Use the Triggers panel to add two timed triggers:

- `update2hPrices` — every **2 h**
- `rolloverDailySheet` — daily at **DAILY_RESET_HOUR:00**

At the end of the script a small sandbox executes `update2hPrices()` once and
logs the result to verify the script runs without errors.
