# quicksheet-fx-ext

Live currency exchange rates for [QuickSheet](https://github.com/cemheren/QuickSheet). Powered by the European Central Bank via [frankfurter.app](https://frankfurter.app) — free, no API key required.

## Install

Type into any cell:

```
ext: github:cemheren/quicksheet-fx-ext
```

## Use

```
fx: USD, EUR, 1, 4
```

Fills 4 rows:

```
1 USD →
EUR 0.89123
ECB 2026-06-13
(blank)
```

### Multiple targets

```
fx: GBP, USD, EUR, JPY, 1, 5
```

```
1 GBP →
USD 1.26543
EUR 1.17234
JPY 198.45
ECB 2026-06-13
```

### Supported currencies

All 30+ currencies published by the ECB: USD, EUR, GBP, JPY, CHF, CAD, AUD, NZD, SEK, NOK, DKK, PLN, CZK, HUF, BGN, RON, TRY, BRL, CNY, HKD, IDR, ILS, INR, KRW, MXN, MYR, PHP, SGD, THB, ZAR.

## Why

Pin exchange rates to your wallpaper — always visible, never a browser tab to hunt for. Pairs naturally with the trader dashboard (`examples/trader-dashboard.csv`).

## Build

Requires .NET 9. Zero NuGet dependencies — only BCL `System.Net.Http`, `System.Text.Json`.

```
dotnet build FxExtension.csproj
```

## Notes

- Data updates once per business day (ECB publishes ~16:00 CET).
- No auth, no rate limits for reasonable use.
- Frankfurter.app is open-source and mirrors ECB reference rates.
- Use `L: <cell>, 1h` in QuickSheet to auto-refresh hourly.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
