# Accounting Extension Suite — Design Doc

Based on market research (2026-05-15). Sources: hledger/ledger/beancount ecosystem,
HN plain text accounting threads, IRS.gov, Frankfurter API, YNAB methodology.

## Key Insight

The plain text accounting community (hledger/ledger/beancount, ~15k combined stars)
has identified the exact pain points. QuickSheet's structural advantage: **live
wallpaper dashboard** — always-visible budget/tax info without opening a browser tab.

## Tier 1 — Build These (highest impact)

### 1. `budget:` — Budget envelope visualizer ⭐⭐⭐⭐⭐
- **Usage:** `budget: Groceries, 800, 623`
- **Output:** Named category with visual fill bar + percentage + surplus/deficit
- **Why:** YNAB has 4M+ users for exactly this mental model. Visual progress bars
  in a terminal screenshot are instantly compelling. ~80 LOC, pure math, no network.
- **Virality:** The "wallpaper as budget dashboard" is a one-sentence pitch.
- **Recipe:**
  ```csv
  Category,Budget,Spent,Tracker
  Groceries,800,623,"budget: Groceries, {B2}, {C2}"
  Software,200,89,"budget: Software, {B3}, {C3}"
  ```

### 2. `qtr:` — Quarterly tax countdown ⭐⭐⭐⭐⭐
- **Usage:** `qtr:` (no params needed, auto-detects next deadline)
- **Output:** Next IRS estimated tax deadline + days remaining + link
- **Why:** Missing quarterly = IRS penalty. Pairs with existing `1099:` extension.
- **IRS deadlines:** Apr 15, Jun 16, Sep 15, Jan 15 (shift for weekends/holidays)
- **Implementation:** DateTime.Now + hardcoded table. ~60 LOC.

### 3. `fx:` — Live currency conversion ⭐⭐⭐⭐
- **Usage:** `fx: 5000, USD, EUR`
- **Output:** Converted amount + rate + source + date
- **API:** Frankfurter `api.frankfurter.dev/v2/rate/USD/EUR` — free, no key, 200 currencies
- **Why:** International freelancers check rates daily before invoicing. ~100 LOC.
- **Cache:** 24h TTL (rates update daily from ECB).

## Tier 2 — Build for Depth

### 4. `rate:` — Freelance hourly rate calculator
- **Usage:** `rate: 120000` (target annual income)
- **Output:** Min viable hourly rate accounting for taxes, benefits, non-billable time
- **Formula:** target ÷ (52 × 40 × 0.7 billable) × 1.55 overhead ≈ required rate
- **Why:** Devs consistently underprice by 30-40%. Educational and surprising output.

### 5. `deduct:` — Schedule C deduction estimator
- **Usage:** `deduct: homeoffice, 250` or `deduct: vehicle, 12000`
- **Categories:** home office ($5/sqft simplified), vehicle (70¢/mi), phone (50%), health
- **Why:** #1 missed deduction for dev freelancers. IRS figures are constants.

### 6. `pl:` — P&L dashboard from cell range
- **Usage:** `pl: {A2::C20}` where col A=label, B=type(income/expense), C=amount
- **Output:** Revenue, expenses, net profit, margin
- **Why:** Live P&L on wallpaper that updates as you add rows.

## Deferred

- `inv:` — Invoice tracker (needs state persistence, more complex)
- `burn:` — Burn rate/runway (needs good range data to exist first)
- `payroll:` — Too jurisdiction-specific
- `receipt:` — File import UX is awkward in current extension model

## The "Freelancer Finance Suite" Vision

Combine existing (`1099:`, `stock:`, `price:`) + new (`budget:`, `qtr:`, `fx:`) into
a cohesive dashboard CSV template. This is the flagship demo: a single CSV file that
turns your desktop wallpaper into a personal finance command center.

## Build Order

1. `budget:` first (highest virality, pure math, no network complexity)
2. `qtr:` second (pairs with existing 1099, trivial)
3. `fx:` third (needs HTTP client, slightly more complex)
4. Tier 2 extensions in later runs
