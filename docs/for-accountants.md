# QuickSheet for accountants & bookkeepers

If you estimate quarterly taxes, track mileage, calculate depreciation, or compute
payroll — **your desktop wallpaper can be the calculator.** No browser tab, no
SaaS subscription. Your numbers stay on your machine.

> Already using Excel or Google Sheets for quick tax estimates? Think "those, but
> always visible on the wallpaper behind your windows, with instant IRS-rate
> extensions you never have to look up."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/accountant-dashboard.csv
```

Edit cells, autosaves to CSV every 5 seconds. Add to startup applications and the
grid is there every workday — *behind* every window, never stealing focus.

## What goes on the wallpaper

Starter sheet at `examples/accountant-dashboard.csv`. Columns:

| Column             | What it shows                                                    | Extension                                                                                  |
|--------------------|------------------------------------------------------------------|--------------------------------------------------------------------------------------------|
| **1099 estimates** | Federal + SE tax on gross income, quarterly payment amount       | [`1099`](https://github.com/Deskworks/quicksheet-1099-ext)                                 |
| **Mileage**        | IRS standard-mileage deduction for business, medical, charity    | [`mileage`](https://github.com/Deskworks/quicksheet-mileage-ext)                           |
| **Depreciation**   | Straight-line and MACRS schedules per IRS Pub 946                | [`depr`](https://github.com/Deskworks/quicksheet-depr-ext)                                 |
| **Break-even**     | Contribution margin, break-even units, profit at volume          | [`margin`](https://github.com/Deskworks/quicksheet-margin-ext)                             |
| **Sales tax**      | State sales tax rates for all 50 states + DC                     | [`tax`](https://github.com/cemheren/quicksheet-salestax-ext)                               |
| **Payroll**        | Federal withholding, Social Security, Medicare per pay period    | [`payroll`](https://github.com/cemheren/quicksheet-payroll-ext)                             |

All extensions install with a single `ext: github:Deskworks/quicksheet-<name>` or
`ext: github:cemheren/quicksheet-<name>` cell. QuickSheet clones the repo, starts
the subprocess, and the prefix is live.

## What the wallpaper IS NOT for

- **Full-form tax filing.** QuickSheet gives instant estimates, not IRS-ready returns.
  Use TurboTax / FreeTaxUSA / your CPA for the actual filing.
- **Double-entry bookkeeping.** If you need a general ledger, QuickBooks or GnuCash
  owns that. The wallpaper is the quick-reference layer, not the system of record.
- **Client billing.** Time tracking and invoicing belong in purpose-built tools.

It's the *reference* layer: "how much should I set aside this quarter?", "what's my
mileage deduction so far?", "when does this asset hit zero?" — answers visible at
a glance while you work.

## Recipe: freelancer quarterly tax dashboard

A freelancer earning variable 1099 income wants to see estimated quarterly payments
and running mileage deductions at a glance.

**Row 1 — Q1 estimate:**
```
1099: 25000
```
Shows federal + SE tax estimate on $25k gross income.

**Row 2 — Q2 estimate:**
```
1099: 32000
```
Updated for cumulative Q2 income.

**Row 3 — Business mileage YTD:**
```
mileage: 4200 2025 business
```
Shows IRS deduction at the current year's standard rate.

**Row 4 — Medical mileage:**
```
mileage: 180 2025 medical
```
Separate rate for medical/moving miles.

**Row 5 — Equipment depreciation:**
```
depr: 3200 5 straight
```
Laptop depreciated over 5 years, straight-line.

**Row 6 — Vehicle (MACRS):**
```
depr: 28000 5 macrs
```
Work vehicle on IRS MACRS 5-year table.

**Row 7 — Break-even analysis:**
```
margin: 5000 45 120
```
Fixed costs $5k, variable $45/unit, sell price $120 — how many units to break even?

**Row 8 — Sales tax check:**
```
tax: CA
```
Quick lookup of California's base rate when quoting a client.

## Recipe: payroll quick-check

Verifying a paycheck or estimating take-home for a new hire:

```
payroll: 85000 biweekly single
```

Shows federal withholding, Social Security (6.2%), and Medicare (1.45%) per pay
period. Useful as a sanity check against the actual pay stub.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula, Synthwave,
Gruvbox, Catppuccin Mocha, Tokyo Night. Some blend better with a financial
spreadsheet aesthetic — try Solarized or Gruvbox for a warm, readable look during
long accounting sessions.

## More

- [60-second tour](tour.md) — all features at a glance
- [Keyboard shortcuts](keyboard-shortcuts.md) — navigate fast
- [Recipes](recipes.md) — dashboard layout patterns
- [Extensions directory](tour.md#extensions) — 85+ extensions available
- [For homelabbers](for-homelab.md) · [For traders](for-traders.md) · [For SREs](for-sre.md) · [For students](for-students.md)
