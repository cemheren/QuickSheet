# QuickSheet for writers

If you're drafting a novel, managing freelance articles, tracking submissions, or building
a research board — **your desktop wallpaper can be the command centre.** No browser tab.
No subscription. Your notes and references live on the wallpaper, always visible behind
every window.

> Already familiar with Scrivener, Notion, or a corkboard of index cards? Think "those,
> but on the actual desktop wallpaper, plain CSV you own forever, and talking to live
> reference APIs through extensions you can write in 50 lines of any language."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/writer-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. Survive reboots — the grid is exactly where
you left it.

## What goes on the wallpaper

The starter sheet at `examples/writer-dashboard.csv` covers five zones you can resize or
swap out:

| Zone                | What it shows                                            | How                                                           |
|---------------------|----------------------------------------------------------|---------------------------------------------------------------|
| **Manuscript**      | Chapter title, word count, status (draft/revise/done)    | Plain cells — edit directly                                   |
| **Submissions**     | Market, genre, date sent, response status                | Plain cells — your submission tracker at a glance             |
| **Research**        | Quick-reference definitions and synonyms                 | [`define`](https://github.com/cemheren/quicksheet-define-ext) / [`thes`](https://github.com/cemheren/quicksheet-thes-ext) |
| **Citations**       | DOI → formatted citation for non-fiction / academic work | [`cite`](https://github.com/cemheren/quicksheet-cite-ext)     |
| **Daily progress**  | Sparkline of word counts over the week                   | `s: B2:B8` sparkline cells                                    |

All extensions install with a single `ext: github:cemheren/quicksheet-<name>` cell.
QuickSheet clones the repo, starts the subprocess, and the prefix is live.

## Why wallpaper > another app

- **Always visible** — no switching to Notion or Scrivener just to check your word count
  target. It's on-screen the moment you unlock.
- **Distraction-free** — it's *behind* every window. When you're in your text editor,
  it's out of the way. When you minimise, it's right there.
- **Survives reboots** — autosaved CSV, opens to the same grid every session.
- **No cloud** — your submission tracker, plot outlines, and research notes are a plain
  file in your folder. No sync. No vendor lock-in.
- **Per-cell hot-swap** — change a single cell, it's saved. No rebuilds, no deploys.

## Recipes

### Word count tracker with sparklines

Track daily output across a week. Put word counts in cells B2–B8, then a sparkline cell:

```
s: B2:B8
```

The sparkline renders an inline bar chart of your writing velocity. Pair it with a Σ
(column sum) to see total words written this week.

### Instant thesaurus while drafting

Stuck on a word? Edit any cell to:

```
thes: mundane
```

The cell fills with synonyms: *ordinary, everyday, banal, prosaic, humdrum*. Copy the one
you want into your manuscript. No browser tab, no context switch.

### Submission tracker

Columns: Market | Genre | Word Count | Date Sent | Status | Days Waiting

The "Days Waiting" column can use a formula or just plain text you update weekly. Sort
rows by status to see what's pending vs. accepted vs. rejected at a glance.

### Research board for non-fiction

Use `cite:` to pull formatted citations:

```
cite: 10.1037/rev0000126
```

Returns the full APA-style citation. Keep a "Sources" zone on your wallpaper so you never
lose track of where that statistic came from.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula, Synthwave,
Gruvbox, Monokai, HotdogStand. Nord and Gruvbox pair well with focused writing
environments.

## Extensions that pair well with writing

- [`define`](https://github.com/cemheren/quicksheet-define-ext) — dictionary lookup.
- [`thes`](https://github.com/cemheren/quicksheet-thes-ext) — thesaurus / synonyms.
- [`cite`](https://github.com/cemheren/quicksheet-cite-ext) — DOI → formatted citation.
- [`wiki`](https://github.com/cemheren/quicksheet-wiki-ext) — Wikipedia summary for quick
  fact-checking.
- [`books`](https://github.com/cemheren/quicksheet-books-ext) — book lookup by ISBN/title.
- [`rss`](https://github.com/cemheren/quicksheet-rss-ext) — subscribe to literary
  magazines, market listings, or writing blogs.

See the [full extension directory](extensions.md).

## Share your setup

Post your writing wallpaper on r/writing, r/selfpublish, r/unixporn, or r/screenwriting.
Open an issue if you want a new prefix officially listed.
