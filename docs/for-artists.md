# QuickSheet for artists & illustrators

If you do commissions, freelance design, or any visual creative work — your
desktop wallpaper is dead space between app windows. **QuickSheet turns it into
a glanceable commission queue, palette swatch, and launcher** that sits *behind*
every window and never steals focus from your canvas.

> Not a project-management tool. A sticky-note layer that's always visible when
> you switch windows, close a tab, or glance at your second monitor.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/artist-dashboard.csv
```

Edit cells live, autosaves every 5s. Add to startup applications and you have a
persistent commission board on every boot.

## What goes on the wallpaper

Starter sheet at `examples/artist-dashboard.csv`. Sections:

| Section              | What it shows                                          | How                                                    |
|----------------------|--------------------------------------------------------|--------------------------------------------------------|
| **Commission queue** | Client, status, due date, rate                         | Plain cells — edit inline                              |
| **Color palette**    | Swatches for your current project                      | `c:red:█` / `c:cyan:█` — named-colour prefix          |
| **Hex references**   | Exact hex codes beside each swatch                     | Plain text cells                                       |
| **App launchers**    | One-click open GIMP, Krita, Blender, Clip Studio       | `r: gimp` / `r: krita` — runnable cell prefix         |
| **Quick links**      | Coolors, Unsplash, PureRef, ArtStation                 | Paste URL → auto-detected hyperlink                   |
| **Income tracker**   | Completed / invoiced / paid this month                 | Plain cells + built-in Σ column sum                   |

## Recipes

### Commission status board

One row per commission. Columns: project name, status, client handle, due date,
rate. Update status as you work — the wallpaper is always there when you tab out
of your drawing app.

Colour-code urgency with the `c:` prefix:
- `c:red:OVERDUE` — past due
- `c:yellow:This week` — due soon
- `c:green:On track` — comfortable

### Palette swatches on your desktop

Use `c:<colour>:█` (full-block character) to paint swatches directly in cells.
QuickSheet supports: red, green, blue, yellow, cyan, magenta, white, gray/grey.

Put hex codes in the adjacent column for clipboard reference. When you start a
new project, swap the row — the palette is always one glance away from your
canvas.

### App launcher row

```
r: gimp
r: krita
r: blender
r: clip-studio-paint
```

Click the cell → app launches. Faster than hunting through a Start menu or dock
when your hands are on a tablet.

### Invoice tracking with column sums

Track monthly income in a simple grid:

```
Project          | Rate  | Status
Portrait         | $180  | Paid
Album cover      | $350  | Invoiced
Emote pack       | $120  | Queued
                 | Σ     |
```

The built-in Σ (column sum) gives you a running total. No spreadsheet formulas
needed — just numbers in a column.

## What the wallpaper IS NOT for

- **Drawing.** Use Krita / Procreate / Clip Studio for your canvas. The wallpaper
  is the *meta* layer beside it.
- **Full project management.** If you need Gantt charts and client portals, use
  Notion / Trello. QuickSheet is the at-a-glance sticky note.
- **Asset storage.** It's a CSV of text cells. Reference your files via
  `r: xdg-open ~/Art/project/` launchers, not embedded images.

## Extensions that pair well

- [`define`](https://github.com/cemheren/quicksheet-define-ext) — quick word
  definitions when writing commission descriptions.
- [`thes`](https://github.com/cemheren/quicksheet-thes-ext) — thesaurus lookups
  for naming pieces.
- [`ping`](https://github.com/cemheren/quicksheet-ping-ext) — verify your
  portfolio site is up.
- [`epoch`](https://github.com/cemheren/quicksheet-epoch-ext) — timestamp
  conversions for deadline math across timezones.
- [Full extension directory](extensions.md).

Install any extension with a single cell: `ext: github:cemheren/quicksheet-<name>`.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula,
Synthwave, Gruvbox, Monokai, HotdogStand. Synthwave or Dracula match the
creative-desk vibe; Light works for daytime streaming.

## What's missing (be honest)

- No embedded image preview in cells. The wallpaper is text-only.
- No Trello/Notion sync. It's intentionally offline-first CSV.
- No time tracking. Use a dedicated tool (Toggl, Clockify) if you bill hourly.

The philosophy: own your data in a plain CSV file, keep it visible, don't let
your commission queue hide behind a browser tab.
