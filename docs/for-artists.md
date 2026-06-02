# QuickSheet for artists & designers

If you draw, paint, model, or design for a living — commissions, freelance,
studio work — your desktop already has reference images scattered across it.
**QuickSheet turns the wallpaper itself into a persistent project board**: client
queue, palette hex codes, deadline countdown, social links, all behind your
Photoshop / Krita / Blender windows. No browser tab, no Notion subscription,
just a CSV that autosaves.

> Not a project-management replacement. A glanceable layer: "what am I working on,
> who's waiting, what's due" — always visible when you minimize everything.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/artist-dashboard.csv
```

Edit cells, autosaves to CSV every 5s. Add to startup applications and the grid
lives behind every window — persistent, zero-focus.

## What goes on the wallpaper

Starter sheet at `examples/artist-dashboard.csv`. Suggested layout:

| Column             | What it shows                                            | How                                                                                    |
|--------------------|----------------------------------------------------------|----------------------------------------------------------------------------------------|
| **Commission queue** | Client name, piece description, status (sketch/lined/done) | Plain cells, colour-code status with `c:green:`, `c:yellow:`, `c:red:`              |
| **Palette strip**  | Hex codes for the current project palette                | `c:#FF6B35: #FF6B35`, `c:#004E89: #004E89` — cells become colour swatches            |
| **Deadlines**      | Days until delivery for each commission                  | [`cntdn`](https://github.com/cemheren/quicksheet-cntdn) extension                    |
| **Reference links** | ArtStation, Pinterest boards, Dribbble, mood boards     | Paste URLs — highlighted, opens on Enter                                               |
| **Income tracker** | Monthly totals with sparkline trend                      | Numbers in column + `s: B1::B12` sparkline                                             |
| **Quick launchers** | Open Krita, Blender, OBS, stream overlay                | `r: krita`, `r: blender`, `r: obs --startstreaming`                                   |
| **Social posts**   | Links to scheduled posts / upload pages                  | `https://twitter.com/compose/tweet`, portfolio URL                                     |

## Recipes

### Colour palette board

Dedicate a row to your active palette. Each cell holds a hex code rendered in
that colour:

```
c:#2D3142: #2D3142
c:#4F5D75: #4F5D75
c:#BFC0C0: #BFC0C0
c:#EF8354: #EF8354
c:#FFFFFF: #FFFFFF
```

Glance at the wallpaper to confirm your palette while painting. Swap codes any
time — the row autosaves.

### Commission pipeline

Track commissions from inquiry to delivery:

```
┌──────────────┬────────────────────┬────────────────┬───────────┐
│ Client       │ Piece              │ Status         │ Due       │
├──────────────┼────────────────────┼────────────────┼───────────┤
│ @user_a      │ Character portrait │ c:green: Done  │ 2026-06-10│
│ @user_b      │ Scene illustration │ c:yellow: WIP  │ 2026-06-18│
│ @user_c      │ Logo redesign      │ c:red: Queued  │ 2026-07-01│
└──────────────┴────────────────────┴────────────────┴───────────┘
```

Status colours are instantly visible at desktop-glance distance.

### Stream / recording launcher strip

One row of runnable cells to start your creative session:

```
r: krita ~/art/current-wip.kra
r: obs --startrecording
r: firefox https://music.youtube.com/playlist?list=...
```

Select all three → Enter → your full setup launches in one keystroke.

## Pair with extensions

| Extension | Use case |
|-----------|----------|
| [`cntdn`](https://github.com/cemheren/quicksheet-cntdn) | Deadline countdown per commission |
| [`news`](https://github.com/Deskworks/quicksheet-news) | ArtStation / Dribbble RSS feed headlines |
| [`pomo`](https://github.com/Deskworks/quicksheet-pomodoro) | Pomodoro timer for focused drawing sessions |
| [`rate`](https://github.com/Deskworks/quicksheet-rate) | Minimum hourly rate calculator against income target |
| [`1099`](https://github.com/cemheren/quicksheet-1099-ext) | Quarterly tax estimates for freelance income |

Install any extension with a single cell: `ext: github:cemheren/quicksheet-<name>`.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula,
Synthwave, Gruvbox, Monokai, HotdogStand. Dracula or Synthwave pair well with
art-heavy desktops; Light keeps things minimal for design work.

Launch with a preset: `dotnet run -c Release --project ExcelConsole.csproj -- --desktop --theme Dracula`.

## What's missing (be honest)

- **No image thumbnails in cells.** The grid is text-only. Your reference images
  live in their own viewer; QuickSheet holds the *links* and *notes*, not previews.
- **No colour-picker integration.** You can't eyedrop from the wallpaper into
  Krita. The palette cells are a reference strip, not an interactive tool.
- **No calendar sync.** Deadlines are manual cells, not synced from Google
  Calendar or iCal.

These are deliberate: the wallpaper is a *glanceable text layer*, not a GUI app.

## More to mix in

- [`fx`](https://github.com/Deskworks/quicksheet-fx) for international commission
  pricing (USD ↔ EUR / GBP / JPY at a glance).
- [`mileage`](https://github.com/cemheren/quicksheet-mileage-ext) if you drive to
  client sites or conventions — IRS standard mileage deduction tracker.
- [Full extension directory](extensions.md).

## Build your own

Need a cell that checks your Etsy / Gumroad / Ko-fi earnings? The QuickSheet
extension protocol is two JSON-lines messages
([extension-protocol.md](extension-protocol.md)) — write a 50-line script that
calls the API and publish it as `quicksheet-<platform>-ext`. The wallpaper picks
it up via `ext: github:user/quicksheet-<platform>-ext`.
