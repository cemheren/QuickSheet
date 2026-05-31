# QuickSheet for accountants & freelancers

If you're tracking contractor income, mileage deductions, depreciation schedules, or
quarterly estimated payments — **your desktop wallpaper can be your at-a-glance tax
dashboard.** No browser tab open to a spreadsheet. No SaaS subscription. Your numbers
stay on your machine in a plain CSV.

> Already using Google Sheets, Excel, or Wave for bookkeeping? Think "a persistent
> dashboard that's always visible behind your windows, fed by extensions you can swap
> in one cell edit."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/accountant-dashboard.csv
```

Edit cells, autosaves to CSV every 5 seconds. Add it to your startup applications and the
grid is there every boot — behind every window, never stealing focus.

## What goes on the wallpaper

The starter sheet at `examples/accountant-dashboard.csv` gives you a frame you can edit
in place:

| Column              | What it shows                                              | Extension used                                                        |
|---------------------|------------------------------------------------------------|-----------------------------------------------------------------------|
| **1099 estimates**  | Quarterly estimated tax payment for self-employment income  | [`1099`](https://github.com/cemheren/quicksheet-1099-ext)             |
| **Mileage**         | IRS standard mileage deduction (business/medical/charity)  | [`mileage`](https://github.com/cemheren/quicksheet-mileage-ext)      |
| **Depreciation**    | Straight-line and MACRS depreciation for assets             | [`depr`](https://github.com/cemheren/quicksheet-depr-ext)            |
| **Margins**         | Break-even point and contribution margin                   | [`margin`](https://github.com/cemheren/quicksheet-margin-ext)        |
| **Payroll**         | Gross-to-net payroll with FICA/Medicare withholding        | [`payroll`](https://github.com/cemheren/quicksheet-payroll-ext)      |
| **Sales tax**       | State + local sales tax lookup                             | [`salestax`](https://github.com/cemheren/quicksheet-salestax-ext)    |

All extensions install with a single `ext: github:cemheren/quicksheet-<name>-ext` cell.
QuickSheet clones the repo, starts the subprocess, and the prefix is live.

## Use cases

### Freelancer quarterly dashboard

You invoice clients all quarter. At a glance, see:

- **Row 1**: `1099: 45000,CA` → shows your estimated federal + state quarterly payment
- **Row 2**: `mileage: 3200,business` → IRS deduction for client site visits
- **Row 3**: Manual cell tracking total invoiced vs. paid this quarter
- **Row 4**: `margin: 150000,95000` → are you above break-even for the year?

### Small business owner

- Track payroll burden: `payroll: 85000,TX` → see net pay and employer FICA cost
- Equipment depreciation: `depr: 12000,5,macrs` → laptop fleet write-off schedule
- Sales tax compliance: `salestax: 99.99,WA,Seattle` → tax amount for each state you sell in

### Tax season prep

Keep a running estimate of your liability visible all year. No more January scramble:

- `1099:` cells update as you change the gross income number
- Mileage cell accumulates as you log trips (edit the miles value after each drive)
- Depreciation cells remind you which assets are fully depreciated this year

## Why wallpaper > spreadsheet tab

- **Always visible** — glance at your tax estimate without switching windows
- **Survives reboots** — autosaved CSV, same grid every boot
- **No SaaS** — your financial data is a local CSV, not on someone else's server
- **Per-cell hot-swap** — change one income figure, the 1099 estimate updates next refresh
- **Plain text audit trail** — CSV diffs in git show exactly what changed and when

## Recipe: quarterly estimated tax tracker

1. One row per income source (client, gig, side project).
2. A `1099:` cell with your projected annual gross → shows quarterly payment.
3. A `mileage:` cell with YTD business miles → shows deduction amount.
4. Manual cells summing deductions to subtract from gross.
5. Add a sparkline (`s:`) across monthly income columns to visualize trends.

## More extensions to mix in

- [`stock`](https://github.com/cemheren/quicksheet-stock-ext) — portfolio positions if you're also investing
- [`price`](https://github.com/cemheren/quicksheet-price-ext) — commodity/crypto prices for cost-basis tracking
- [`mortgage`](https://github.com/cemheren/quicksheet-mortgage-ext) — amortization schedules for real estate investors

See the [full extension directory](extensions.md).

## Honesty

- **Not accounting software.** QuickSheet is a quick-glance dashboard, not a double-entry
  ledger. It won't generate Schedule C or file anything for you.
- **Tax rates may lag.** Extensions use hard-coded IRS rate tables that need annual updates.
  Always verify against irs.gov before filing.
- **No bank integration.** You enter numbers manually or pipe them from your own scripts.
  There's no Plaid, no API keys, no OAuth.
