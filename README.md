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

- `update2hPrices` — every **2 h**
- `rolloverDailySheet` — daily at **DAILY_RESET_HOUR:00**

Example formulas for the five Δ columns (row 2):

```text
E2: =IFERROR((B2-B3)/B3, "")
F2: =IFERROR((B2-B5)/B5, "")
G2: =IFERROR((B2-B9)/B9, "")
H2: =IFERROR((B2-B13)/B13, "")
I2: =IFERROR((B2-B14)/B14, "")
```

**表头必须 9 列，脚本自动补空列；若想手动添加公式请从 E2 开始向右写 Δ2h ~ Δ24h。**
