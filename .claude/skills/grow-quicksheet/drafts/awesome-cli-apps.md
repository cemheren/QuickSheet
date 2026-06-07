# awesome-cli-apps submission — draft (2026-06-06)

> User submits the PR manually. Do not push to agarrharr/awesome-cli-apps.

## ⚠️ Gate: Requires >20 GitHub stars

The [contributing guidelines](https://github.com/agarrharr/awesome-cli-apps/blob/master/contributing.md)
explicitly require: "Have more than 20 stars (if it is hosted on GitHub)." QuickSheet
currently has 0 stars. **Do not submit until the repo crosses 20 stars.**

## Target

Repo: https://github.com/agarrharr/awesome-cli-apps (~20k★, actively maintained)
File: `readme.md`
Section: **Data Manipulation** (top-level, alongside visidata and sc-im — the two
closest comparisons already on the list).

## One-line entry (matches list's existing style)

The list uses: `- [APP_NAME](LINK) - DESCRIPTION.` — description starts with a
capital, ends with a period, no "CLI" or "terminal" (redundant given the list name).

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet) - Spreadsheet that doubles as your desktop wallpaper with runnable cells and an extension system.
```

Shorter variant:

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet) - Spreadsheet with a desktop-wallpaper mode, runnable command cells, and CSV persistence.
```

Place at the **bottom** of the Data Manipulation section (per contributing guide: "Add
the app at the bottom of the relevant category").

## PR title

```
Add QuickSheet
```

(Per guidelines: "title it simply `Add APP_NAME`" — nothing else.)

## PR body

```markdown
**QuickSheet** is an interactive spreadsheet (.NET 9, MIT) that runs as a TUI or
embeds itself as an interactive desktop wallpaper (Windows WorkerW, Linux X11).

**Why Data Manipulation (alongside visidata, sc-im):**
- Same category of tool — interactive spreadsheet in the terminal
- Differentiator: desktop-wallpaper embedding + cell prefixes that run shell commands,
  stream live subprocess output, render sparklines, and install extensions from git repos

**Checklist (from contributing.md):**
- [x] Does one thing well (spreadsheet with CSV persistence)
- [x] Free and open source (MIT)
- [x] Easy to install (`dotnet tool install` or single binary from Releases)
- [x] Well documented (README + docs/ + tour.md)
- [x] Older than 90 days
- [x] More than 20 stars

Repo: https://github.com/cemheren/QuickSheet
```

## Notes for the submitter

- **One PR per app.** Don't bundle with other changes.
- **Use the PR template** the repo provides (the contributing.md is explicit: "Use the
  provided pull request template. Failure to follow this point means the PR will be
  closed without being looked at.").
- The description avoids "CLI" and "terminal" per their style guide (it's redundant
  on a CLI-apps list).
- Placement in "Data Manipulation" is correct — visidata ("Spreadsheet multitool for
  data discovery") and sc-im ("Vim-like spreadsheet calculator") are there. QuickSheet
  is the same genre: interactive spreadsheet.
- If the maintainer asks "why not Productivity?" — answer: it's a data tool first
  (CSV in, CSV out). The wallpaper mode is a display trick, not a productivity feature.
- The list recently moved from `aharris88` to `agarrharr` as primary maintainer.
  PRs are being merged regularly (latest merges within past weeks).

## After submission (expected impact)

- Inbound link from a ~20k-star list. This is one of the highest-value awesome-lists
  for terminal tools — it's linked from Sindre's master awesome-list.
- Long-tail: consistent 10–50 stars/month from search traffic + list browsers.
- The list is mirrored by trackawesomelist.com and other aggregators, compounding reach.

## Ranking among awesome-* submissions

| List                                   | Star count | Min-stars req | Hit-prob | Status |
|----------------------------------------|------------|---------------|----------|--------|
| luong-komorebi/Awesome-Linux-Software  | ~22k       | none          | high     | draft ready (`drafts/awesome-linux-software.md`) |
| **agarrharr/awesome-cli-apps**         | **~20k**   | **>20 stars** | **high** | **this draft (GATED)** |
| 0PandaDEV/awesome-windows              | ~2.4k      | none noted    | medium   | draft ready (`drafts/awesome-windows.md`) |
| secretGeek/AwesomeCSV                  | ~926       | none          | high     | draft ready (`drafts/awesome-csv.md`) |
| awesome-csharp (uhub)                  | smaller    | unknown       | high     | `drafts/awesome-lists.md` §2 |
| awesome-dotnet (quozd)                 | medium     | unknown       | high     | `drafts/awesome-lists.md` §2 |
| awesome-tuis (rothgar)                 | medium     | unknown       | marginal | `drafts/awesome-lists.md` §1 |

## Submission priority

1. **Now (0 stars):** AwesomeCSV, Awesome-Linux-Software, awesome-windows — no star gate.
2. **After 20 stars:** awesome-cli-apps (this draft), awesome-dotnet, awesome-csharp.
3. **After demo GIF:** awesome-tuis (leads with screenshots; wallpaper demo needed).
