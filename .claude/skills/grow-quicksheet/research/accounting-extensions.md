# Accounting extensions: gap analysis + concrete next builds

Research run 2026-05-15 for [QuickSheet#15](https://github.com/cemheren/QuickSheet/issues/15).

## Summary

- Three accounting/finance extensions exist: `1099-ext` (US SE-tax), `mortgage-ext` (loan amortization), `qtr` (likely quarterly tax). Plus tangentially: `budget`, `fx`, `price-ext` (crypto), `stock-ext`, `rate`.
- Real gaps for a US-centric self-employed / small-LLC audience: **sales-tax-by-state lookup**, **mileage/per-diem deduction calculator**, **invoice-aging / DSO**, **payroll/W-2 net-pay estimator**, **depreciation (MACRS/straight-line)**, **break-even / margin**, **CapEx ROI / NPV-IRR**.
- Best ROI (low complexity, high "I needed this yesterday" factor for the user persona that stars dev tools): **mileage**, **depreciation**, **break-even/margin**. All pure-math, zero-network, deterministic — they slot into the same shape as `1099-ext` and `mortgage-ext`.
- Decline: anything that depends on a paid API (avalara, taxjar), anything that requires storing PII, anything needing a real bookkeeping ledger (out of cell-prefix scope).

## Existing vs. gaps

| Vertical                       | Status                  | Notes                                                                |
|--------------------------------|-------------------------|----------------------------------------------------------------------|
| Self-employment tax (US)       | Have: `1099-ext`        | SE rate hard-coded, no state.                                        |
| Loan amortization              | Have: `mortgage-ext`    | Fixed-rate, monthly. ARM/biweekly are future.                        |
| Quarterly estimated tax        | Have: `qtr` (assumed)   | Verify scope before adding overlap.                                  |
| Budget tracking                | Have: `budget` (assumed)| Likely simple total-vs-cap.                                          |
| FX                             | Have: `fx`              | Spot rate.                                                           |
| Sales tax by state             | **Gap**                 | Static table per-state, lookup is one-shot, no API needed.           |
| Mileage / per-diem deduction   | **Gap**                 | IRS std rate × miles. One number per year (2024: $0.67/mi business). |
| Invoice aging / DSO            | **Gap**                 | Days-sales-outstanding from invoice list. Needs row-range input.     |
| Net pay (W-2)                  | **Gap**                 | Gross → fed/state/FICA → net. State table similar to sales-tax.      |
| Depreciation                   | **Gap**                 | MACRS table or straight-line. Both pure math.                        |
| Break-even / margin            | **Gap**                 | Fixed + variable + price → break-even units. One formula.            |
| CapEx ROI (NPV/IRR)            | **Gap**                 | Standard NPV / IRR over cashflow vector. Math only.                  |

## Top 3 concrete builds (queue)

### 1. `quicksheet-mileage-ext` — IRS mileage deduction

Cell input:
```
mileage: 1250, business
```

Output (3 cells):
```
1,250 mi business
× $0.67/mi (IRS 2024)
= $837.50 deduction
```

Modes: `business` ($0.67), `medical` ($0.21), `charity` ($0.14). Rates table hard-coded per tax year, fall back to latest with a "year unknown" cell. Pure math. ~80 LOC. Zero deps.

Repo description: "IRS standard-mileage deduction calculator for QuickSheet — business/medical/charity rates."

### 2. `quicksheet-margin-ext` — Break-even & contribution-margin

Cell input:
```
margin: 50000, 12, 30
```
(fixed cost, variable per unit, sell price)

Output (4 cells):
```
Contribution margin: $18 / unit
Break-even units: 2,778
Break-even revenue: $83,340
Margin ratio: 60.0%
```

Add `margin: 50000, 12, 30, units=5000` mode to also compute profit at a target volume. ~60 LOC.

### 3. `quicksheet-depr-ext` — Straight-line + MACRS depreciation

Cell input:
```
depr: 10000, 5, straight
```
or
```
depr: 10000, 5, macrs
```

Output:
- Straight-line: one row per year, equal $.
- MACRS: lookup the 5/7/15/20-year half-year-convention table from IRS Pub 946.

~120 LOC (MACRS tables are the bulk). Pure math, no network.

## Out of scope / declined

- **Real sales-tax-by-zip**: state-level is one static table (50 entries) but city/county/special-district lookups require taxjar/avalara. Decline — static 50-state extension would be misleading (real sales tax is hyper-local).
- **Real-time payroll**: legitimate W-4-aware calculation needs current IRS Pub 15-T tables + state-by-state withholding. Worth doing in a follow-up, but tables are large enough it'd be a >300 LOC ext.
- **Bookkeeping / ledger**: out of cell-prefix scope. Belongs in a separate app, not a wallpaper cell.

## Implications for QuickSheet (action queue)

1. **Next Bucket F run: build `quicksheet-mileage-ext`** — simplest, smallest, near-universally useful for any US freelancer.
2. **Run after: build `quicksheet-margin-ext`** — one formula, four output cells, the kind of thing that screenshots well.
3. **Run after: build `quicksheet-depr-ext`** — slightly bigger (MACRS table) but rounds out the small-business trio.
4. **Comment on issue #15** linking this brief and listing the three queued builds. Leave issue open until at least one of the three is shipped.

After those three ship, issue #15 can close.
