# Persona 3 — Accountants & bookkeepers

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- ~1.4M US accountants + ~1.5M bookkeepers. Spreadsheets are literally the job; CSV-first is a feature, not a limitation.
- Already partially served by extensions: `1099`, `qtr`, `mileage`, `margin`, `depr`. Need to (a) tell accountants the cluster exists, and (b) fill the obvious next gaps.
- Highest-leverage gaps: `payroll:` (FICA/Medicare/state withholding), `sales-tax:` (per-state rate lookup), `gl:` (general-ledger trial balance from CSV), `cogs:` (FIFO/LIFO/weighted-avg), `reconcile:` (bank-vs-book diff).
- Pain: practice mgmt (QBO, Xero) is required for clients but every accountant maintains a *parallel* personal scratch Excel for the actual thinking. QuickSheet replaces that scratch layer, not QBO.
- Where to find them: r/Accounting, r/Bookkeeping, r/taxpros, AICPA's CPA.com community, Going Concern blog, Accounting Today.

## 1. Profile

- US: ~1.4M accountants/auditors + ~1.5M bookkeeping/accounting/auditing clerks (BLS 2024). Solo + small-firm CPAs ≈ 50k EAs + 100k+ CPAs.
- Day-shape: client books reconciliation, payroll runs, monthly close, tax season spikes (Jan–Apr + extension Sep–Oct).
- Tax season behavior: 12-hour days, dozens of clients, ambient deadline anxiety. **Anything that surfaces a deadline on the desktop wins.**
- Pays attention to: accuracy, auditability, "your data stays here," tools that don't break QuickBooks integration.

## 2. Current toolchain

| Tool                 | Used for                                | Pain                                                          |
|----------------------|------------------------------------------|----------------------------------------------------------------|
| QuickBooks Online    | Client books of record                   | Required; slow; tab-only; not glanceable.                       |
| Xero                 | Same (different firms)                   | Same                                                            |
| Excel (personal)     | "Scratch" — reconciliations, what-ifs    | Files scatter across desktop; no live data; no automation.     |
| Drake / UltraTax     | Tax prep                                 | Desktop apps; not designed for ambient info.                    |
| QBO Accountant tools | Client list view                          | Login dance every morning.                                      |
| Sticky notes / paper | Deadline tracking                        | Lost. Tax extensions missed are malpractice-adjacent.          |

Unmet need: **a desktop layer that shows (a) which clients have an action due this week, (b) live FICA/sales-tax/withholding lookups, (c) a personal scratch grid that survives reboots — without yet another SaaS to subscribe to.**

## 3. Pain points QuickSheet could touch

1. **Tax / filing deadline visibility.** Quarterly estimates, 1099 deadlines, payroll deposits, state-specific dates.
2. **Lookup-heavy work.** "What's the FICA cap this year?" "What's CA sales tax?" "What's the per-diem for Boston?"
3. **Reconciliation diffs.** Bank CSV vs book CSV — line up, find mismatches.
4. **Per-client status board.** 50 clients, who's overdue, who needs a draft sent.
5. **Personal time tracking** (overlap with lawyers' `bill:`).

## 4. Candidate extensions / features (ranked)

| Rank | Extension                          | Cost | Hit probability | Why                                                              |
|------|------------------------------------|------|------------------|-------------------------------------------------------------------|
| 1    | `payroll:` FICA + Medicare + state | low  | high             | Pure tables. Updates yearly. Top "what is the cap" lookup.        |
| 2    | `sales-tax:` per-state rate        | low  | high             | Avalara has free CSV; bake table. Hits every retail bookkeeper.   |
| 3    | `reconcile:` bank vs book diff     | med  | high             | Local CSV diff; flag unmatched lines. The monthly-close core.     |
| 4    | `gl:` trial balance from CSV       | med  | med-high         | Group debits/credits by account; show out-of-balance.            |
| 5    | `cogs:` FIFO/LIFO/weighted-avg     | low  | med              | Pure math; pairs with inventory CSVs.                              |
| 6    | `perdiem:` GSA per-diem lookup     | low  | med              | GSA has free JSON. Auditors live here during travel season.       |
| 7    | `irspub:` IRS publication snippets | high | med              | Scrape, cache, search. Big lift; defer.                            |
| 8    | `1099`, `qtr`, `mileage`, `depr`, `margin` | shipped | — | Already exist. Need to *tell this audience* they exist.           |

QuickSheet *features* worth queuing for this persona:

- **`audit:` mode** — every cell edit appends `(timestamp, oldval, newval)` to a sidecar `.audit.csv`. Lawyers + accountants + medical care. Bucket E.
- **Locked-cell prefix** (`!: ...`) — read-only cells, useful for "this is the closing balance, do not edit." Bucket E, ~30 LOC.
- **Multi-sheet** support is the obvious limitation accountants will hit. Out of scope for now (would break CSV-first hard rule unless implemented as multi-CSV directory).

## 5. Where they hang out

- **Reddit:** r/Accounting (~600k), r/Bookkeeping, r/taxpros, r/CPA, r/smallbusiness (the *clients* of solo bookkeepers).
- **Forums/communities:** AICPA Connect, NATP forum, Intuit ProConnect community, AccountingWEB, Going Concern (Tumblr-era blog but still active).
- **Podcasts:** Cloud Accounting Podcast (Blake Oliver), Earmark CPE, Accounting Today's podcast.
- **Conferences:** AICPA Engage (Vegas, June), Scaling New Heights, QuickBooks Connect.
- **Newsletters:** CPA Trendlines, Going Concern, Accounting Today, Tax Pro Today.
- **People to be visible to:** Blake Oliver (Cloud Accounting), Caleb Newquist (Going Concern alum), Heather Smith (cloud accounting writer/Australia).

## 6. Discoverability hooks

This audience treats software the way lawyers do: skeptically, with a high "where does my data live" sensitivity.

Headlines that would land:

- "I built a free, open-source desktop dashboard for solo bookkeepers"
- "QuickBooks tabs are not a dashboard. This is."
- "A 50-line extension to look up FICA caps without opening IRS.gov"

Lead screenshot: a grid with (a) one row of `qtr:` countdown timers, (b) one row of `1099:` calculations for sample clients, (c) one cell of `sales-tax: CA` showing 7.25 %, (d) `mileage:` totals. **One image = "this software speaks your language."**

## 7. Implications (queue these)

1. **Build `payroll:` extension** (Bucket F). FICA/Medicare/state withholding lookup. Pure tables, updates yearly, hits every accountant. Pair with `sales-tax:` (Bucket F).
2. **Bucket A: write `docs/for-accountants.md`** — a curated "if you're a bookkeeper, here are the extensions and a starter CSV." Audience-specific landing page. Link from README.
3. **Bucket C draft: r/Accounting and Cloud Accounting Podcast outreach.** Save in `drafts/accountants-launch.md`. Lead with screenshot above. Cite zero-deps, MIT, CSV-first.
4. **Bucket E: `audit:` mode** for cell-edit history. Trust signal for this persona + lawyers + clinicians.

Cross-link:
- [accounting-extensions.md](../accounting-extensions.md) — earlier gap analysis; this paper expands on the audience side.
- [lawyers.md](lawyers.md) — overlap on audit-trail / CSV-trust angle and `bill:` timer feature.
