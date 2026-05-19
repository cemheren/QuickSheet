# quicksheet-mortgage-ext

Fixed-rate mortgage payment calculator for [QuickSheet](https://github.com/cemheren/QuickSheet). No network, no API key, just the amortization formula.

## Install

Type into any cell:

```
ext: github:Deskworks/quicksheet-mortgage-ext
```

## Use

```
mort: 500000, 6.5, 30, 1, 4
```

Fills 4 rows:

```
$500,000 @ 6.5% / 30yr
monthly: $3,160.34
total interest: $637,724
total paid: $1,137,724
```

Parameters are `principal, annual_rate_percent, years`. Rate accepts decimals (`6.875`).

## Why

Mortgage and loan numbers are exactly the kind of thing you want to compare side-by-side on a wallpaper-pinned spreadsheet — 30 vs 15 year, fixed vs adjustable, different rate scenarios. Drop a few `mort:` cells next to each other, and you have a live calculator with no spreadsheet formulas to write.

Works for any fixed-rate fully-amortizing loan: car loans, student loans, personal loans.

## Build

Requires .NET 9. Zero NuGet dependencies — only `System.Globalization` and `System.Text.Json` from BCL.

```
dotnet build MortgageExtension.csproj
```

## Notes

- Pure math. No network, no cache, no state.
- US-style amortization formula (`P × r / (1 − (1+r)^-n)`).
- Rate is the annual nominal rate; the monthly rate is `r/12`.
- Does not include taxes, insurance, PMI, or HOA — those are policy decisions, not math.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
