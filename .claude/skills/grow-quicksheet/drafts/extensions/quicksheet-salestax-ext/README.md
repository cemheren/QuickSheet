# quicksheet-salestax-ext

US state sales tax rate lookup for [QuickSheet](https://github.com/cemheren/QuickSheet).

Instantly check any state's sales tax rate — state-level and combined (state + average local) — right on your desktop. Optionally pass a dollar amount to see the tax and total.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-salestax-ext
```

## Usage

| Cell content | Output |
|---|---|
| `tax: CA` | California (CA) — State: 7.25%, Combined: 8.85%, On $100: $8.85 |
| `tax: TX, 250` | Texas (TX) — Tax on $250: $20.50, Total: $270.50 |
| `tax: Oregon` | Oregon (OR) — No sales tax ✓ |
| `tax: New York` | New York (NY) — State: 4%, Combined: 8.53% |

### Parameters

```
tax: <state> [, amount]
```

- **state** — Two-letter abbreviation (CA, TX, NY) or full name (California, Texas).
- **amount** *(optional)* — Dollar amount to calculate tax on.

## Data

Rates are combined state + average local rates from Tax Foundation 2024 data. All 50 states + DC included. States with no general sales tax (OR, MT, NH, DE, AK) show `0%` or note local-only rates.

## Build

```bash
dotnet build SalesTaxExtension.csproj
```

Requires .NET 9. Zero NuGet dependencies.

## License

MIT
