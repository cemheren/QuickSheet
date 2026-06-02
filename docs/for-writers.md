# QuickSheet for writers & academics

If you context-switch between drafts, source tabs, citation managers, and
scattered notes — **your desktop wallpaper can be the calm orientation layer
above all of that.** No browser tab. No cloud subscription. Always visible
behind every window.

> Already use Scrivener, Obsidian, or Zotero? QuickSheet doesn't replace them.
> It's the glance-down surface where you track *where you are* without opening
> anything.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/writer-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. Your grid is exactly where you left
it after a reboot.

## What goes on the wallpaper

The starter sheet at `examples/writer-dashboard.csv` covers four zones:

| Zone               | What it shows                                  | How                                                              |
|--------------------|------------------------------------------------|------------------------------------------------------------------|
| **Manuscript map** | Chapter titles + status (draft / revised / done) | Plain cells — edit directly                                    |
| **Research queue** | Open questions, DOIs to read, notes            | Plain cells + `cite:` for formatted citations                    |
| **Word tools**     | Definitions, synonyms at a glance              | [`def`](https://github.com/Deskworks/quicksheet-define-ext) · [`thes`](https://github.com/cemheren/quicksheet-thes-ext) |
| **arXiv feed**     | Latest papers in your field                    | [`arxiv`](https://github.com/Deskworks/quicksheet-arxiv)         |

All extensions install with a single `ext: github:user/repo` cell.

## Manuscript tracker

Plain cells are enough for a chapter outline:

```
Chapter,Status,Word count,Notes
1 – The hook,🟢 Done,3200,Needs tighter opening line
2 – Backstory,🟡 Revising,4100,Cut flashback to 2 pages
3 – Inciting incident,🟡 Revising,2800,
4 – Rising action,⬜ Draft,1900,Merge with ch 5?
5 – Midpoint,⬜ Not started,,
```

Ctrl+F to jump to any chapter. Ctrl+B to sort by status. This is always visible
below your editor window — no app-switching needed.

## Citation workflow

Install the citation extension:

```
ext: github:cemheren/quicksheet-cite-ext
```

Then paste a DOI into any cell:

```
cite: 10.1038/s41586-023-06185-3
```

Output fills the row: authors, year, title, venue — formatted and ready to copy
into your reference list.

For arXiv papers:

```
ext: github:Deskworks/quicksheet-arxiv
arxiv: 2301.07041
```

Returns title, authors, abstract snippet. Keep a column of paper IDs as your
reading queue.

## Word tools on the wallpaper

### Dictionary

```
ext: github:Deskworks/quicksheet-define-ext
def: ephemeral
```

Returns part of speech + definition. Useful when you're deciding between two
words without opening a browser tab.

### Thesaurus

```
ext: github:cemheren/quicksheet-thes-ext
thes: ephemeral
```

Fills rows with synonyms (transient, fleeting, momentary…). Pick the one that
fits your sentence.

## Research notes that don't disappear

Use a column as a running research log:

```
Question,Source,Answer
What year was X discovered?,cite: 10.xxxx/yyyy,1923 — Smith et al.
How does Y relate to Z?,,TODO: read Jones 2019
Reviewer asked about sample size,,Check supplemental table S3
```

This stays on your wallpaper while you write. No separate app, no lost sticky
note, no buried Notion page.

## Runnable commands for writers

Prefix a cell with `r:` to make it a one-click launcher:

```
r: code ~/manuscripts/thesis/
r: xdg-open ~/references/annotated.pdf
r: wc -w ~/manuscripts/thesis/chapter*.md
```

Press Enter on the cell to run. Your manuscript folder, key PDFs, and word-count
checks — always one keystroke away.

## Who this is for

- **Academics**: lit review tracking, citation queue, deadline awareness.
- **Non-fiction / Substack writers**: chapter map, source log, word-count targets.
- **Fiction writers**: scene outline, character reference, revision status.
- **Journalists**: source contacts, story status board, quick fact-lookups.

The common thread: you context-switch constantly, and the wallpaper is the one
surface that's *always there* without demanding focus.

## Install

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop
```

Or grab a binary from [Releases](https://github.com/cemheren/QuickSheet/releases)
(when available).

Requirements: .NET 9 SDK, Linux (X11) or Windows 10+.
