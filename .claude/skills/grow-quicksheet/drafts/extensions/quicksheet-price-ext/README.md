# quicksheet-price-ext

Live crypto price quotes (CoinGecko) for [QuickSheet](https://github.com/cemheren/QuickSheet).

## Install

Type into any cell:

```
ext: github:cemheren/quicksheet-price-ext
```

## Use

```
price: btc, 1, 2
```

Fills 2 rows starting at the cell:

```
bitcoin $94,213
▲ 1.42% 24h
```

Accepts common tickers (`btc`, `eth`, `sol`, `doge`, `ada`, `xrp`, `dot`, `matic`, `link`, `ltc`, `bch`, `avax`, `atom`, `arb`, `op`) or any CoinGecko coin id directly:

```
price: ethereum
price: solana
price: chainlink
```

## Why

If your wallpaper is a QuickSheet, a row of `price:` cells is the simplest "always-on portfolio glance" you can build. No browser tab, no app — just a CSV of ticker symbols.

Combined with `L: <cell>, 1m` you can have the values refresh on whatever interval you want.

## Build

Requires .NET 9. Zero NuGet dependencies — just `System.Net.Http` and `System.Text.Json` from BCL.

```
dotnet build PriceExtension.csproj
```

## Notes

- Uses the free CoinGecko API. No key required.
- Responses are cached for 60 seconds to avoid rate-limit churn.
- Prices in USD only for now.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
