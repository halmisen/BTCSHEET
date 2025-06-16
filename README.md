# BTCSHEET

This repository contains sample Google Apps Script code to fetch cryptocurrency prices from Coinbase every 2 hours.

The `coinbase_2h.gs` script provides two main functions:

- `update2hPrices()` fetches the latest 2-hour candle close for BTC-USD, ETH-USD and SOL-USD from Coinbase Advanced Trade and appends it to the `Data` sheet in the active spreadsheet if the timestamp is new.
- `initHistory(limit)` populates the `Data` sheet with the most recent 2-hour candles (100 by default).

The data is retrieved using the public Coinbase API endpoint:

```
https://api.exchange.coinbase.com/products/BTC-USD/candles?granularity=7200&limit=1
```

See the official documentation for details: <https://docs.cloud.coinbase.com/exchange/reference/exchangerestapi_getproductcandles>.

The public API allows about 3 requests per second, so this script only queries a few products at low frequency.

After deploying the script, create a trigger in Google Sheets to run `update2hPrices()` every 2 hours for continuous updates.

At the end of the script a small sandbox executes `update2hPrices()` once and logs the result to verify the script runs without errors.
