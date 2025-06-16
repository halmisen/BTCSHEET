# BTCSHEET

This repository contains sample Google Apps Script code to fetch cryptocurrency prices from Binance every 2 hours.

The `binance_2h.gs` script provides two main functions:

- `update2hPrices()` fetches the latest 2-hour kline for BTCUSDT, ETHUSDT and SOLUSDT and appends it to the `Data` sheet in the active spreadsheet if the timestamp is new.
- `initHistory(limit)` populates the `Data` sheet with historical data (100 klines by default).

At the end of the script a small sandbox executes `update2hPrices()` once and logs the result to verify the script runs without errors.
