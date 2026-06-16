# QuickSheet for freelancers

You track clients, invoices, hours, expenses, and quarterly taxes — usually across
three apps and a dozen browser tabs. **Your wallpaper can be the single glanceable
layer** showing what's due, what's overdue, and how your year is tracking. No SaaS
subscription, no cloud sync, your numbers stay on your machine in a plain CSV.

> Not a replacement for your accounting software. A complement: the always-visible
> summary that tells you "invoice #47 is 12 days overdue" and "Q3 estimated tax is
> due in 9 days" without opening anything.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/freelancer-dashboard.csv
```

Edit cells live, autosaves to CSV every 5s. Add to startup applications and the
dashboard is there every morning — behind every window, never stealing focus.

## What goes on the wallpaper

Starter sheet at `examples/freelancer-dashboard.csv`. Sections:

| Section              | What it shows                                            | Extension / feature                                                                        |
|----------------------|----------------------------------------------------------|--------------------------------------------------------------------------------------------|
| **Invoice tracker**  | Client, amount, date sent, days outstanding              | Built-in CSV — colour overdue rows with `c:red:`                                           |
| **Hours this week**  | Daily hours logged per client, weekly total              | Built-in `Σ` column sum                                                                    |
| **Quarterly tax**    | Days until next IRS estimated-tax deadline              | [`qtr`](https://github.com/cemheren/quicksheet-cal-ext) (countdown)                        |
| **SE tax estimate**  | Federal + state self-employment tax from YTD gross      | [`1099`](https://github.com/cemheren/quicksheet-1099-ext)                                  |
| **Mileage log**     | Trip date, miles, deduction at IRS rate                 | [`mileage`](https://github.com/cemheren/quicksheet-mileage-ext)                            |
| **Upcoming dates**   | Next 7 days from your calendar feed                     | [`cal`](https://github.com/cemheren/quicksheet-cal-ext) (ICS feed)                         |
| **Expense running total** | YTD business expenses by category                | Built-in `Σ` column sum + sparkline `s: B2::B13`                                          |

All extensions install with a single `ext: github:cemheren/quicksheet-<name>` cell.

## Recipes

### Invoice aging row

One row per outstanding invoice. Columns:

| A (Client) | B (Invoice #) | C (Amount) | D (Date sent) | E (Status) |
|------------|---------------|------------|----------------|------------|
| Acme Corp  | INV-047       | $3,200     | 2026-05-28     | `c:red: 18 days` |
| Widgets LLC| INV-048       | $1,500     | 2026-06-10     | `c:green: 5 days` |

Update manually when paid, or use an `i: curl` cell to poll your invoicing API.

### Weekly hours heatmap

```
Mon  Tue  Wed  Thu  Fri  Sat  Sun  Σ
 6    8    7    8    5    0    0   34
```

Each cell is just a number. The Σ row auto-sums. Add a sparkline across 52 weeks
to see your utilisation trend: `s: B2::B53`.

### Quarterly estimated tax countdown

Place this in any cell:

```
ext: github:cemheren/quicksheet-cal-ext
```

Then in another cell: `cal: next-quarter` shows days remaining until the next
estimated payment (April 15, June 15, Sept 15, Jan 15).

## Pair with the 1099 cluster

If you file as sole proprietor or single-member LLC:

- [`1099`](https://github.com/cemheren/quicksheet-1099-ext) — estimates your
  quarterly payment from YTD gross income.
- [`mileage`](https://github.com/cemheren/quicksheet-mileage-ext) — tracks
  business miles at the current IRS rate ($0.70/mile for 2026).
- [`salestax`](https://github.com/cemheren/quicksheet-salestax-ext) — looks up
  state + local sales tax for invoicing clients in other jurisdictions.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula,
Synthwave, Gruvbox, Monokai, HotdogStand. Light or Solarized keep things
readable during work hours; Dracula when you're burning midnight oil.

## What's missing (be honest)

- **No time-tracking integration.** You can't start/stop a timer from a cell
  (timers are explicitly out of scope). Log hours manually or paste from Toggl.
- **No Stripe/PayPal webhook listener.** QuickSheet is offline-first. Poll
  with `i: curl` if your invoicing platform has an API.
- **No multi-currency invoicing.** Use the [`fx`](https://github.com/cemheren/quicksheet-fx-ext)
  extension to show conversion rates, but invoice amounts are plain text.

The wallpaper is your *at-a-glance status board*, not your full accounting stack.

## Get started

```bash
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/freelancer-dashboard.csv
```

Edit the CSV to match your actual clients. Autosave keeps it current. The grid is
always there when you Alt+Tab out of your editor or browser.
