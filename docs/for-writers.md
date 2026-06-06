# QuickSheet for writers & researchers

If you're drafting a novel, a thesis, a blog series, or a research paper — **your desktop
wallpaper can be your reference board.** Character names, citation DOIs, word definitions,
synonym alternatives, chapter outlines — always visible, never buried in a tab.

> Already familiar with Scrivener, Obsidian, or Zotero? Think "those reference panels, but
> on the actual wallpaper, zero cloud, and talking to free APIs through extensions you can
> write in 50 lines of any language."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/writer-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. The grid sits behind every window — glance at
your notes without Alt-Tabbing away from your editor.

## What goes on the wallpaper

The starter sheet at `examples/writer-dashboard.csv` gives you four zones:

| Zone              | What it shows                                              | How                                                         |
|-------------------|------------------------------------------------------------|-------------------------------------------------------------|
| **Outline**       | Chapter titles, status, word-count targets                 | Plain cells — edit directly                                 |
| **Research**      | DOI citations formatted inline, arXiv paper lookups        | [`cite`](https://github.com/Deskworks/quicksheet-cite-ext) · [`arxiv`](https://github.com/Deskworks/quicksheet-arxiv) |
| **Language**      | Definitions, synonyms — right there on the desktop         | [`def`](https://github.com/Deskworks/quicksheet-define-ext) · [`thes`](https://github.com/Deskworks/quicksheet-thes-ext) |
| **Quick refs**    | Character names, timeline, recurring motifs                | Plain cells — your scratch pad                              |

All extensions install with a single `ext: github:Deskworks/quicksheet-<name>` cell —
QuickSheet clones the repo and the prefix is live instantly.

## Look up a word without leaving your editor

Need a definition while you write? It's on the wallpaper:

```
ext: github:Deskworks/quicksheet-define-ext
---
def: ephemeral
def: liminal
def: palimpsest
```

Each cell shows the word's definition inline. No browser, no dictionary app, no context
switch.

## Synonyms at a glance

Stuck on a word? The thesaurus extension shows alternatives:

```
ext: github:Deskworks/quicksheet-thes-ext
---
thes: beautiful
thes: walk
thes: said
```

Keep a column of "words I overuse" with live synonym suggestions beside your draft.

## Citation lookups for academics

Paste a DOI and get a formatted citation without opening Google Scholar:

```
ext: github:Deskworks/quicksheet-cite-ext
---
cite: 10.1145/3359591.3359737
cite: 10.1038/s41586-020-2649-2
```

Returns: Author(s), Year, Title, Venue. Copy into your LaTeX `\cite{}` or bibliography
manager.

## arXiv paper search

Looking up recent papers in your field:

```
ext: github:Deskworks/quicksheet-arxiv
---
arxiv: 2301.07041
arxiv: attention is all you need
```

Returns title, authors, year, and abstract snippet. Keep a "reading list" column on your
wallpaper.

## Chapter outline as a living document

Plain cells are all you need:

```
Chapter,Status,Word Count,Notes
1 - The Arrival,🟢 Done,4200,Hook established
2 - The Market,🟡 Drafting,2800,Needs more sensory detail
3 - The Letter,⬜ Outline only,0,Plot twist reveal
4 - The Return,⬜ Not started,0,Mirror chapter 1 structure
```

Ctrl+B to sort by status. Ctrl+F to search across all cells. The grid is always there when
you look up from your manuscript.

## Worldbuilding & character reference

Dedicate a few rows to the details you keep forgetting:

```
Name,Role,Trait,First Appearance
Maren,Protagonist,Left-handed; scar on chin,Ch 1
Isa,Mentor,Speaks in questions,Ch 3
The Compact,Organization,Seven members; meets at solstice,Ch 2
```

No more flipping through a separate "bible" document — it's on the wallpaper behind your
word processor.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula, Synthwave,
Gruvbox, Monokai, HotdogStand. Nord or Solarized pair well with long writing sessions.

## Why this over Notion / Obsidian / Scrivener?

| Feature                         | QuickSheet                          | Notion/Obsidian          |
|---------------------------------|-------------------------------------|--------------------------|
| Always visible                  | ✅ On the wallpaper, behind windows | ❌ Must open/switch app   |
| Zero cloud dependency           | ✅ CSV on disk, no account          | ⚠️ Varies                |
| Live API lookups (define, cite) | ✅ Extensions, no browser           | ❌ Plugin ecosystem       |
| Privacy                         | ✅ Local-only, no telemetry         | ⚠️ Varies                |
| Cost                            | ✅ Free, MIT                        | ⚠️ Freemium              |

QuickSheet doesn't replace your writing app — it's the **reference board beside it.**
