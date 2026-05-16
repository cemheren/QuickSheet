# Persona 2 — Lawyers (solo + small-firm)

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- Solo + 2–10-attorney firms ~400k+ in the US. Underserved by modern desktop tooling, locked into legacy practice-management SaaS.
- Spreadsheets are *native*: billable hours, conflict checks, matter lists. CSV-first matches the audit-trail mindset.
- High-leverage extensions: `bill:` (6-min increments), `dock:` (court docket pull), `stat:` (statute lookup), `cite:` (Bluebook formatter — distinct from academic `cite`), `clock:` (matter timer).
- Pain: practice-management software (Clio, MyCase) is $50–100/mo per seat and feels like SaaS bloat. QuickSheet's "your data is a CSV" is a *feature*, not a limitation.
- Where to find them: r/Lawyertalk, r/LawFirm, r/solopractice, Lawyerist podcast & community, ABA TECHSHOW, /r/lawyers.

## 1. Profile

- ~1.3M licensed US attorneys; ~75 % work in firms of <20 attorneys; ~50 % in firms of <10. Source: ABA Profile of the Legal Profession 2024.
- Solo practitioners alone ≈ 400k. They write their own systems.
- Day-shape: meetings + drafting + discovery review + billing capture. *Billing capture is the productivity black hole.*
- Pays attention to: tools that respect data ownership (audit, retention), don't phone home, and look professional next to a client on a screen-share.

## 2. Current toolchain

| Tool                              | Used for                                | Pain                                                                  |
|-----------------------------------|------------------------------------------|------------------------------------------------------------------------|
| Clio / MyCase / PracticePanther    | Billing, matter mgmt, document storage   | $50–$120/seat/month; "another SaaS to log into."                       |
| Excel (constantly)                 | Conflict checks, matter lists, time logs | Files on desktop everywhere; version drift across paralegals.          |
| Outlook / Gmail                    | Communication + dated artifacts          | Time spent re-entering "what did I do for client X yesterday."         |
| Westlaw / Lexis                    | Research                                 | Per-search billing anxiety; tab graveyard.                              |
| PACER                              | Federal court docket access              | $0.10/page; tooling around it is from 1998.                            |
| Sticky notes                       | "Don't forget to call back"              | Lost; not audit-trail-able.                                            |

Unmet need: **a desktop layer that captures small atomic facts (time entries, statute citations, matter notes) without launching an app, then exports clean CSV** when it's time to send to billing software or paralegal.

## 3. Pain points QuickSheet could touch

1. **Billing capture in 6-min increments** while drafting. The hardest, most-money problem.
2. **Conflict-check list** open on a second monitor while taking a call: search a CSV of every client/adverse party.
3. **Court calendar / docket countdown** — deadlines are malpractice if missed.
4. **Statute & rule snippet pinning** — recall "what does 28 U.S.C. § 1331 say" without opening Westlaw.
5. **Bluebook citation formatter** — distinct from academic `cite:`; uses *signal + reporter + court + year* format.

## 4. Candidate extensions / features (ranked)

| Rank | Extension                            | Cost | Hit probability | Why                                                                |
|------|--------------------------------------|------|------------------|---------------------------------------------------------------------|
| 1    | `bill:` — 6-min increment timer      | low  | high             | Type matter ID → cell starts ticking; auto-rounds to .1h. CSV out. |
| 2    | `dock:` — federal court deadline calc| med  | high             | FRCP day-counting (skip weekends/holidays). Hand-coded table.       |
| 3    | `cite:` — Bluebook formatter         | low  | med-high         | Lookup of reporter abbreviations; rename existing `cite:` → `doi:`. |
| 4    | `stat:` — USC / state statute lookup | med  | med              | Cornell LII has free, scrapable text. Cache aggressively.           |
| 5    | `conflict:` — fuzzy match local CSV  | low  | high             | Pure-local fuzzy match against `clients.csv` next to the QuickSheet.|
| 6    | `pacer:` — docket pull               | high | med              | Real auth + paid. Defer until users ask.                            |
| 7    | `caselaw:` — Caselaw Access Project  | low  | med              | Free, no auth. Returns case name, court, year, snippet.             |

QuickSheet *features* worth queuing for this persona:

- **Per-cell timer prefix** (`t:` perhaps) — Bucket E. Cell auto-increments while focused; right-click stop. Output is `.h` (decimal hour). Foundational for `bill:`.
- **CSV → Markdown table export already exists** (`--export-md`) — perfect for pasting into a "weekly billing summary" email. Should be marketed *to this persona* explicitly.
- **"Audit log" mode** — every edit writes a timestamp to a side log file. Lawyers + accountants + medical all care.

## 5. Where they hang out

- **Reddit:** r/Lawyertalk (private, ~120k), r/LawFirm (~50k), r/solopractice, r/legaltech (~10k but high signal).
- **Podcasts:** Lawyerist Podcast, Above the Law, Legal Talk Network. Lawyerist also has an online community ("Lawyerist Lab").
- **Conferences:** ABA TECHSHOW (Chicago, every spring), Clio Cloud Conference, ALA conference.
- **Newsletters:** Bob Ambrogi's "LawSites," 3 Geeks and a Law Blog, Lawyerist Insider.
- **Discords/Slacks:** smaller; the Lawyerist Lab Slack is the closest to a "hangout."
- **People to be visible to:** Bob Ambrogi (LawSites), Carolyn Elefant (MyShingle), Sam Glover (Lawyerist), Nicole Black (MyCase legal tech writer). They review tools genuinely.

## 6. Discoverability hooks

This audience is *intensely* allergic to startup-bro language. Wins are framed as:

- "I built a desktop tool that captures billable time without making me open another app"
- "Open-source, your data is just CSV, runs offline, no SaaS to subscribe to"
- "For lawyers who already use Excel for everything"

Headlines that would land in r/Lawyertalk or LawSites:

- "I built a desktop spreadsheet for tracking billable hours — local, free, no SaaS"
- "My matter list now sits on my desktop wallpaper. Here's the build."
- "A Bluebook citation formatter you can paste into any cell"

Lead with the **billing-capture screenshot** (matter ID + ticking timer + day's hours summed). That image alone does the selling — every solo lawyer has cried over forgotten time entries.

## 7. Implications (queue these)

1. **Build `bill:` extension** (Bucket F). 6-min auto-rounding timer cell. Highest-leverage on the list. Pairs naturally with `--export-md` for end-of-day billing summary.
2. **Rename `cite:` → `doi:` and free up `cite:` for Bluebook** (Bucket A breaking-rename — needs deprecation path: keep `cite:` aliasing to `doi:` for 1 release, log a one-line warning). Bluebook formatter is the higher-frequency use.
3. **Bucket C draft: r/Lawyertalk post + LawSites email pitch to Bob Ambrogi.** Save in `drafts/lawyers-launch.md`. User submits. Lead with billing-capture screenshot, end with "MIT-licensed, your data stays on your machine."
4. **Bucket B: add `legal`, `lawtech`, `billing` topics** to QuickSheet repo once a legal extension exists (no false advertising before then).

Cross-link: [accountants.md](accountants.md) (next persona) will overlap heavily on the audit/CSV-trust angle.
