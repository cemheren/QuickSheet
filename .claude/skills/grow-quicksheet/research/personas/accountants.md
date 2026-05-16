# Persona 3 — Accountants & bookkeepers (deadline-storm desktop)

Status: **done** (v2, desktop-mode focus) · Last revised: 2026-05-16

## TL;DR (5 lines)

- ~1.4M US accountants + ~1.5M bookkeepers. Most are single- or dual-monitor; QuickBooks and tax-prep apps eat the foreground; **wallpaper is the only never-closed surface during tax season.**
- They are not TUI users. Pitch is "your desktop wallpaper shows every client's deadline state and your unfiled-return count, before QuickBooks even loads."
- Highest-leverage wallpaper cells: per-client status row (Q-ready / awaiting docs / overdue), tax-deadline countdowns, FICA/sales-tax/per-diem lookups, daily $-billable, audit-trail strip.
- Where to seed: r/Accounting, r/Bookkeeping, r/taxpros, Cloud Accounting Podcast, AccountingWEB, Going Concern. *Plus* r/desktops with the "tax-season desktop" angle.
- Direct competition for wallpaper slot: zero. QBO is browser, Drake is desktop app, Excel is windowed. Nobody is selling a wallpaper layer to this audience.

## 1. Profile

- US: ~1.4M accountants/auditors + ~1.5M bookkeeping/auditing clerks (BLS 2024). Solo + small-firm CPAs ≈ 50k EAs + 100k+ CPAs.
- Hardware: 1–2 monitors. The good ones already use a second monitor for "client status." That is our seat.
- Day-shape: client books reconciliation, payroll runs, monthly close, **tax-season storm Jan–Apr + extension Sep–Oct**. During the storm, the laptop never sleeps; deadline anxiety is constant.
- Buying brain: "Does this respect audit-trail? Does my data stay on my machine? Will it survive when I move to a new laptop?" CSV-first answers all three.

## 2. Why their desktop is wasted

What lives there now:

- A static wallpaper or firm logo.
- Folder graveyard: `Smith 2024`, `Jones LLC payroll`, `Q3 estimates`.
- Maybe QBO Accountant in a perma-tab they have to focus to see.

What QuickSheet replaces: the *intent* of opening QBO Accountant every hour to "see where I am with my clients." With the wallpaper, that view is **already painted on the desktop** every time they switch out of Excel. Zero clicks, zero refresh.

## 3. Glanceable data they actually want behind windows

1. **Per-client status row** — N cells, one per client. Colour: green = up-to-date, yellow = needs docs, red = overdue. Click to drill in.
2. **Deadline strip** — quarterly estimates, 1099 deadlines, payroll deposits, state filings. "12d to Q3 estimated tax."
3. **Today's $-billable** — single big cell summing matter timers (overlap with lawyers).
4. **Unfiled returns counter** — single cell, "8 unfiled" in red.
5. **Lookup row** — three cells: `payroll: FICA-cap`, `sales-tax: CA`, `perdiem: Boston`. Always there, no Google.
6. **Reconciliation diff cell** — `reconcile: bank.csv, books.csv` shows count of unmatched lines.
7. **Audit log corner** — last 5 cell edits with timestamps. Reassurance.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                | Type | Cost | Hit prob | Why                                                                       |
|------|-------------------------------------|------|------|----------|----------------------------------------------------------------------------|
| 1    | **Value-driven cell colour**        | feat | low  | high     | Same feature needed for SREs + lawyers. Trinity unlock.                    |
| 2    | **Cell staleness dimming**          | feat | low  | high     | "Is this number still current?" answered visually.                         |
| 3    | **`audit:` sidecar log mode**       | feat | med  | high     | Every cell edit logged to `.audit.csv`. Audit-trail trust signal.          |
| 4    | `payroll:` FICA/Medicare/state      | ext  | low  | high     | Pure tables. Daily lookup. Updates yearly.                                  |
| 5    | `sales-tax:` per-state rate         | ext  | low  | high     | Avalara CSV bake. Hits every retail bookkeeper.                            |
| 6    | `reconcile:` bank vs book CSV diff  | ext  | med  | high     | Monthly-close core. Output is a count + first mismatch line.                |
| 7    | `gl:` trial balance from CSV         | ext  | med  | med-high | Sum debits/credits by account, flag imbalance. Pairs with `reconcile:`.    |
| 8    | `perdiem:` GSA per-diem lookup       | ext  | low  | med      | Free JSON; auditors live in it during travel season.                       |
| 9    | `cogs:` FIFO/LIFO/weighted-avg       | ext  | low  | med      | Pure math; pairs with inventory CSV.                                       |
| 10   | `irspub:` IRS pub snippets            | ext  | high | med      | Scrape + cache. Big lift; defer.                                            |
| —    | `1099`, `qtr`, `mileage`, `depr`, `margin` | shipped | — | —    | Already shipped. Need a *desktop screenshot* of them in use.              |

## 5. Where they hang out (desktop-relevant first)

- **r/desktops, r/macsetups** — tax-season-desktop screenshot is a novel post here. Title: "My accountant tax-season desktop: client statuses + IRS deadlines live."
- **r/Accounting** (~600k), **r/Bookkeeping**, **r/taxpros**, **r/CPA** — profession side.
- **Podcasts:** Cloud Accounting Podcast (Blake Oliver), Earmark CPE, Accounting Today.
- **Forums:** AICPA Connect, NATP forum, Intuit ProConnect community, AccountingWEB, Going Concern.
- **Conferences:** AICPA Engage (June, Vegas), Scaling New Heights, QuickBooks Connect.
- **People to be visible to:** Blake Oliver (Cloud Accounting), Heather Smith, Nicole Black, Bob Ambrogi (legaltech overlap).

## 6. Discoverability hooks

Hero image = **the tax-season laptop**. Window: QuickBooks Online in a browser. Wallpaper behind/around it: a 12-client row colour-coded, a `qtr: 9d to Q3 estimate` countdown in red, a `payroll: FICA-cap 2026 = $176,100` lookup cell.

Headlines that land:

- "I put every client's status on my desktop wallpaper. Tax season got quieter."
- "A free, open-source wallpaper dashboard for solo bookkeepers"
- "QuickBooks tabs are not a dashboard. This is."

Avoid: terminal screenshots, "TUI," dev jargon. Lead with **the deadline you almost missed**, not the protocol.

## 7. Implications (queue these)

1. **Bucket E: ship value-driven cell colour + cell staleness dimming + `audit:` sidecar.** These three features together make the wallpaper *credible* for any audit-trail-sensitive persona (accountants, lawyers, clinicians). High build leverage — one Bucket E run can ship two of the three.
2. **Build `payroll:` + `sales-tax:` extensions** (Bucket F). Both are pure-table lookups, no auth, hit every accountant daily.
3. **Bucket A: `docs/for-accountants.md`** — audience-specific landing page with the hero screenshot above + extension cluster + "your data stays here" pitch. Link from README.
4. **Bucket C draft: r/Accounting + Cloud Accounting Podcast pitch.** `drafts/accountants-launch.md`. Save only after the colour + staleness + audit features ship.

Cross-link:
- [accounting-extensions.md](../accounting-extensions.md) — earlier gap analysis (Bucket F focused).
- [lawyers.md](lawyers.md) — overlap on timer + audit-trail value-prop.
