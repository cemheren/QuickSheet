# quicksheet-stock-ext

Stock ticker quotes (Stooq) for [QuickSheet](https://github.com/cemheren/QuickSheet). Free, no API key.

## Install

```
ext: github:Deskworks/quicksheet-stock-ext
```

## Use

```
stock: AAPL, 1, 3
```

Fills 3 rows:

```
AAPL.US $189.85
▲ 0.18% (+0.35)
as of 2026-05-13
```

US tickers default to the `.us` Stooq suffix. For other markets, pass the suffix explicitly:

```
stock: bp.uk        ← BP on LSE
stock: 7203.jp      ← Toyota on TSE
stock: spy.us       ← S&P 500 ETF
```

## Why

A column of `stock:` cells next to your tickers is a watchlist that lives on the wallpaper — no Yahoo tab, no broker app.

Combined with `L: <cell>, 30m` you can keep it gently fresh through the trading day.

## Build

Requires .NET 9. Zero NuGet dependencies — only BCL.

```
dotnet build StockExtension.csproj
```

## Notes

- Uses Stooq's free CSV endpoint (`https://stooq.com/q/l/`).
- Stooq's data is end-of-day for most exchanges; expect 15-minute or longer delay during the trading day.
- Responses cached in-memory for 5 minutes.
- Volume column is fetched but not displayed in the default 3-row layout.

## License

MIT
