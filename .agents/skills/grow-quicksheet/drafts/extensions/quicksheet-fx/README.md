# quicksheet-fx

Live currency conversion extension for [QuickSheet](https://github.com/cemheren/QuickSheet) — the terminal spreadsheet that runs as your desktop wallpaper.

Uses the [Frankfurter API](https://frankfurter.dev) (European Central Bank rates). **No API key required.** 200+ currencies. Rates update daily.

## Install

Type in any QuickSheet cell:
```
ext: github:cemheren/quicksheet-fx
```

## Usage

```
fx: 1000, USD, EUR           → Convert $1,000 to EUR
fx: 5000, USD, EUR, GBP, CAD → Convert to multiple currencies
fx: EUR, JPY                  → Show rate for 1 EUR → JPY
```

### Output example

```
💱 1,000.00 USD │ Rate   │ Converted
→ EUR           │ 0.8551 │ 855.14 EUR
→ GBP           │ 0.7413 │ 741.30 GBP
→ CAD           │ 1.3714 │ 1,371.40 CAD
ECB · 2026-05-15
```

## Features

- **200+ currencies** — all currencies tracked by the European Central Bank
- **No API key** — completely free, no signup, no rate limits for normal use
- **1-hour cache** — avoids redundant API calls (ECB rates update once per day)
- **Multi-target** — convert to multiple currencies in one cell
- **Auto-format** — thousands separators, 2 decimal places for amounts, 4 for rates

## Perfect for

- **International freelancers** — check rates before invoicing clients in different currencies
- **Digital nomads** — quick budget conversions on your desktop wallpaper
- **Finance dashboards** — combine with `1099:`, `budget:`, `qtr:` for a complete freelancer finance suite

## Part of the Freelancer Finance Suite

| Extension | Prefix | What it does |
|-----------|--------|-------------|
| [quicksheet-1099-ext](https://github.com/cemheren/quicksheet-1099-ext) | `1099:` | Self-employment tax estimation |
| [quicksheet-budget](https://github.com/cemheren/quicksheet-budget) | `budget:` | Budget envelope tracking with visual bars |
| [quicksheet-qtr](https://github.com/cemheren/quicksheet-qtr) | `qtr:` | IRS quarterly tax deadline countdown |
| **quicksheet-fx** | `fx:` | **Currency conversion (this extension)** |

## Requirements

- [QuickSheet](https://github.com/cemheren/QuickSheet) with extension support
- .NET 9 SDK
- Internet connection (for API calls)

## License

MIT
