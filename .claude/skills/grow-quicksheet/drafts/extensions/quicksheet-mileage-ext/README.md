# quicksheet-mileage-ext

US IRS standard-mileage deduction calculator for [QuickSheet](https://github.com/cemheren/QuickSheet). Pure math, no network. Not tax advice.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-mileage-ext
```

## Use

Type one of these in a cell:

| Cell contents                       | Output                                                  |
|-------------------------------------|---------------------------------------------------------|
| `mileage: 1250, business`           | 1,250 mi × $0.700/mi (2025) = $875.00 deduction         |
| `mileage: 800, medical`             | 800 mi × $0.210/mi (2025) = $168.00 deduction           |
| `mileage: 500, charity`             | 500 mi × $0.140/mi (2025) = $70.00 deduction            |
| `mileage: 1250, business, 2024`     | Uses 2024 rates ($0.670/mi business)                    |

Modes: `business` / `biz`, `medical` / `med` / `moving`, `charity` / `char`. Defaults to `business`.

Years: 2021–2025. Unknown year falls back to the latest known.

## Example

Type `mileage: 1250, business` in cell **A1**:

|     | A                                |
|-----|----------------------------------|
| **1** | 1,250 mi business                |
| **2** | × $0.700/mi (IRS 2025)           |
| **3** | = $875.00 deduction              |

## IRS rates (source: irs.gov/tax-professionals/standard-mileage-rates)

| Year | Business | Medical / Moving | Charity |
|------|----------|------------------|---------|
| 2025 | $0.700   | $0.210           | $0.140  |
| 2024 | $0.670   | $0.210           | $0.140  |
| 2023 | $0.655   | $0.220           | $0.140  |
| 2022 | $0.625*  | $0.220           | $0.140  |
| 2021 | $0.560   | $0.160           | $0.140  |

\* 2022 had a mid-year change (Jan-Jun $0.585, Jul-Dec $0.625 business). This extension uses the H2 rate as the default.

## Not tax advice

Standard mileage is one of two methods (the other is actual expenses). Eligibility, recordkeeping, and limits depend on your specific situation. Consult a CPA.

## License

MIT
