# Adjacent-project pitch teardown — what desktop-widget tools lead with (2026-05-17)

## TL;DR (5 lines)

- Compared README/landing-page openers for **Rainmeter**, **Conky**, **Übersicht**, and **GeekTool** (the four most-comparable wallpaper-data tools).
- All four lead with **category + benefit**, not metaphor. Conky's "free, light-weight system monitor for X" and Übersicht's "keep an eye on what's happening" are the cleanest.
- QuickSheet's current opener **"Your desktop is a spreadsheet"** is metaphor-first and *more distinctive than any peer's*. Keep the metaphor; bolt a one-line benefit underneath for the readers who skim past taglines.
- QuickSheet's *unique* selling points vs all four peers: (a) cells can run arbitrary shell commands (none of the peers do this), (b) CSV-as-persistence (peers all use proprietary configs), (c) cross-platform Windows + Linux with the same data file (Rainmeter Win-only, Conky X11-only, Übersicht macOS-only, GeekTool macOS-only).
- Three concrete README hero edits queued at the bottom; ship as a separate Bucket A PR.

## Data points

### Rainmeter (~5k★ on GitHub, ~10M users)

- Tagline: *"Rainmeter is a desktop customization tool for Windows."*
- Repo positioning: category-first, vague. The README spends more lines on build infrastructure and code-signing than on "what does it do."
- Marketing site (rainmeter.net) was 403 for WebFetch; based on prior knowledge: leads with skin gallery screenshots, "your CPU/RAM/clock on your desktop."
- Primitive: **skins** — pre-built widget bundles users download or write in a custom INI-like config language.
- Lock-in: configs are Rainmeter-specific.
- Audience-fit signal: "customization" is the keyword. Their hero is *not* benefit-first.

### Conky (~10k★)

- Tagline: *"Conky is a free, light-weight system monitor for X, that displays any kind of information on your desktop."*
- Strengths in the pitch:
  - Word "light-weight" before anything else. They know that's the differentiator.
  - "X, Wayland, and other things, too" — explicit about platform.
  - **300+ built-in objects** — they quantify the surface area.
- Primitive: a config file with `${cpu}`, `${memory_bar 4}`, etc. — interpolation-language inside a paint-the-wallpaper engine.
- Lock-in: the conkyrc syntax.
- Hero-pattern lesson: **lead with the constraint you're proud of** (lightweight), **then the breadth** (300 objects).

### Übersicht (macOS, ~5k★)

- Tagline: *"Keep an eye on what's happening on your machine and in the world."*
- Strengths in the pitch:
  - Benefit-first. No mention of how it works in the tagline.
  - "and in the world" — explicitly opens the door beyond system stats.
- Primitives: JavaScript/JSX widgets with `command`, `refreshFrequency`, `render`, `updateState`, `initialState` hooks. Plus a `run()` shell helper.
- Lock-in: widgets are JS/JSX, not portable.
- Hero-pattern lesson: **benefit-first tagline beats mechanism-first.** Reader knows "ambient view" before they know about JSX.

### GeekTool (macOS, closed-source, dormant)

- WebFetch failed on both candidate URLs. Based on prior knowledge:
- Three primitives: **Shell** (paint stdout on the desktop), **File** (tail/paint a file), **Image** (paint a URL/image).
- Pitch: position-then-paint. Drag a "geeklet" onto the desktop, set its command, see output.
- Lesson: the *primitive count* matters less than how the user thinks about it. Three things, each obvious.

## Where QuickSheet sits

Current README hero (verbatim):

```
# QuickSheet

**Your desktop is a spreadsheet.**

QuickSheet replaces your wallpaper with a transparent, interactive grid. Pin notes,
launch apps, paste links, track numbers — all without opening a window. The idea:
keep something lightweight always present in the background, instead of a static
wallpaper you never interact with.
```

**Diagnosis vs peers:**

| Dimension                     | Rainmeter | Conky | Übersicht | GeekTool | QuickSheet                                   |
|-------------------------------|-----------|-------|-----------|----------|----------------------------------------------|
| Tagline style                  | category | category+constraint | benefit | category | **metaphor** (unique) |
| Quantifies primitive count    | no       | yes (300) | no | no | not in hero (40+ extensions in badges only) |
| Names persistence format      | no       | no | no | no | **no** (CSV is *the* hook; should be in hero) |
| Names cross-platform reach    | no (Win only) | yes (X+Wayland) | no (macOS only) | no (macOS) | partial ("Windows + Linux" only in badges) |
| Names interactive cells       | no       | no | no | no | **partial** — "interactive grid" stated, "type a shell command" not |

**Verdict:** the metaphor "Your desktop is a spreadsheet" is the strongest single line of any tool in this set. Keep it. The follow-up paragraph is fine but it buries three things peers would put in the hero:

1. **CSV persistence** ("Your data is a file in your folder, not a config language").
2. **Runnable cells** ("Type `r: code .` in a cell. Press Enter. The IDE opens.").
3. **Cross-platform with the same file** ("Same CSV on your Windows laptop and your Linux desktop").

These are QuickSheet's unique differentiators vs all four peers — they're not in the hero.

## Implications (queue these)

### 1. Bucket A: README hero edit (small, additive)

Right under the existing 2-sentence paragraph, add a single sentence — *not a feature list*, not a section, just one more line — that names what makes QuickSheet not-Rainmeter-not-Conky:

```
The data is a CSV file. Cells can run shell commands. Same file on Windows or Linux.
```

That's 13 words, factual, and answers the "wait, how is this different from Rainmeter?" question a HN/r/unixporn reader will ask in the first 3 seconds.

Ship as a separate Bucket A PR after the user has had a chance to react to *this* brief.

### 2. Bucket A: badge for cell-prefixes count

The existing "Extensions 40+" badge is good. Add one more:

```
![Cell prefixes](https://img.shields.io/badge/cell_prefixes-6-orange)
```

(`r:`, `i:`, `s:`, `c:`, `L:`, `ext:` — six in-grid primitives.) Mirrors Conky's "300 objects" quantification at the badge level. Honest, small, immediately scannable.

### 3. Bucket C: Show HN draft revision

Existing `drafts/showhn.md` should lead with the metaphor + the 13-word differentiator line, not the mechanics. A 1-line edit to that draft.

### 4. Do NOT add a feature comparison table to the README

Comparison tables read as defensive. The hero edit + badge are enough. If a commenter asks "vs Rainmeter," the answer goes in a reply, not the README.
