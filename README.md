# BTCSHEET

This repository contains sample Google Apps Script code to fetch cryptocurrency prices from Coinbase every 2 hours.

The `trade_manager.gs` script exposes several utilities:

- `update2hPrices()` rebuilds the `Data` sheet with the latest **13** two hour candles for BTC-USD, ETH-USD and SOL-USD and automatically refreshes the change summary table beneath the data.
- `fetchLatest2hCandles(product, limit)` builds recent 2h candles from 1h data.
- `rolloverDailySheet()` copies the current `Data` sheet to a new sheet named by date and then refreshes `Data` for the new day.
- `backfillHistory(start, end)` downloads historical two hour candles between two dates and stores them in a sheet named `History_<start>_to_<end>`.
- `refreshLatestChanges()` rebuilds the summary table below the main data showing price differences for the last row compared to the previous 2h, 4h, 12h and 24h rows.
- `initLedger()` creates the `Ledger` sheet and validation rules.
- `recomputeLedger()` calculates running positions and floating P&L for all trades.

The data is retrieved using the public Coinbase API endpoint:

```
https://api.exchange.coinbase.com/products/BTC-USD/candles?granularity=3600&limit=2
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

- `update2hPrices` — every **2 h**
- `rolloverDailySheet` — daily at **DAILY_RESET_HOUR:00**

**表头只需 4 列（Timestamp 及各币种价格），脚本会在表格下方自动创建 Δ2h ~ Δ24h 汇总表。**

The `refreshLatestChanges()` function rebuilds this summary table and fills the
cells with values such as `+97.2 (0.09%)`.
