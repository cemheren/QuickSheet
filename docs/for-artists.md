# QuickSheet for artists & creatives

If you draw, paint, sculpt, do photography, or freelance in any visual medium —
**your desktop wallpaper can be your studio dashboard.** Commission queue,
colour palettes, word banks, project deadlines, reference links, all visible the
moment you minimise Photoshop / Krita / Blender. No browser tabs, no Notion, no
subscription.

> Not a replacement for your canvas. A complement: the layer that answers
> "what am I working on, who's waiting, what's the hex code" while you paint.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/artist-dashboard.csv
```

Edit cells live, autosaves to CSV every 5s. Add to startup and the grid is
there every session — behind every window, never stealing focus.

## What goes on the wallpaper

Starter sheet at `examples/artist-dashboard.csv`. Zones:

| Zone              | What it shows                                               | How                                                                                        |
|-------------------|-------------------------------------------------------------|--------------------------------------------------------------------------------------------|
| **Commission Q**  | Client name, subject, status, deadline                      | Plain cells — type freely, colour with `c:green:` / `c:yellow:` / `c:red:`                 |
| **Colour palette**| Hex codes + named swatches for current project              | [`color`](https://github.com/Deskworks/quicksheet-color) — `color: #FF6B35` shows name     |
| **Word bank**     | Synonyms / mood words for naming pieces                     | [`thes`](https://github.com/Deskworks/quicksheet-thes-ext) — `thes: serene`                |
| **Definitions**   | Quick lookup while titling work                             | [`define`](https://github.com/Deskworks/quicksheet-define-ext) — `define: ephemeral`       |
| **Deadlines**     | Days until next delivery / convention / gallery submission  | [`cal`](https://github.com/Deskworks/quicksheet-cal) — `cal: 2025-03-15` → days remaining  |
| **Income track**  | YTD commission income, quarterly tax estimate               | [`1099`](https://github.com/Deskworks/quicksheet-1099-ext) — `1099: 4200`                  |
| **Inspiration**   | Links to references, moodboards, Artstation pages           | Plain URLs — auto-detected as clickable hyperlinks                                          |

## What the wallpaper IS NOT for

- **Drawing.** It's a text grid, not a canvas. Your art app owns the pixels.
- **File management.** Use your OS file browser for PSD/TIFF/blend files.
- **Client communication.** Email/Discord still handles that. The wallpaper is
  your private overview, not a shared workspace.

It's the *meta* layer: who owes you money, what's due next week, what colour
is `#2E4057` called, and what word means "slightly blue." Glanceable context.

## Recipe: commission tracker

1. Column A: client name.
2. Column B: piece subject / description.
3. Column C: status — `c:green: Done`, `c:yellow: WIP`, `c:red: Overdue`.
4. Column D: deadline — `cal: 2025-04-01` shows countdown.
5. Column E: price.
6. Bottom row: `Σ` column sum for total outstanding revenue.

All visible while you paint. When a piece ships, flip the cell to green and
the wallpaper updates in 5 seconds.

## Recipe: colour reference strip

Keep a row of hex codes for your current palette:

```
#FF6B35  |  color: #FF6B35  |  #2E4057  |  color: #2E4057  |  #1B998B  |  color: #1B998B
```

The `color` extension resolves each hex to its nearest named colour and
complementary. Useful when you need to communicate colours to a client in words.

## Pair with the freelance cluster

Most artists freelance. Pair your studio dashboard with:

- [`1099`](https://github.com/Deskworks/quicksheet-1099-ext) — quarterly
  estimated tax from YTD gross.
- [`rate`](https://github.com/Deskworks/quicksheet-rate) — minimum viable
  hourly rate against your target annual income.
- [`mileage`](https://github.com/Deskworks/quicksheet-mileage-ext) — IRS
  standard mileage deduction for convention / gallery travel.
- [`qtr`](https://github.com/Deskworks/quicksheet-qtr) — days until next
  estimated-tax deadline.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula,
Synthwave, Gruvbox, Monokai, HotdogStand. Nord or Gruvbox pair well with
creative workflows — muted backgrounds that don't compete with your art.

## What's missing (be honest)

- No image previews in cells. The grid is text-only; keep reference images in
  your file browser or a moodboard app.
- No calendar sync. Deadlines are manual cells, not pulled from Google Calendar.
- No invoice generation. Track income here, produce invoices elsewhere.

These are intentional: the wallpaper is a glanceable scratchpad, not a project
management suite.

## More to mix in

- [`define`](https://github.com/Deskworks/quicksheet-define-ext) for quick
  word lookups while naming a series.
- [`news`](https://github.com/Deskworks/quicksheet-news) subscribed to an art
  industry RSS feed (ArtNet, Colossal, etc.).
- [`epoch`](https://github.com/Deskworks/quicksheet-epoch) for timestamping
  when a commission was accepted.
- [Full extension directory](extensions.md).

## Build your own

Have a niche tool you wish existed? The QuickSheet extension protocol is two
JSON-lines messages ([extension-protocol.md](extension-protocol.md)) — write a
50-line script in Python / Go / .NET and publish it as
`quicksheet-<name>-ext`. The wallpaper picks it up via
`ext: github:user/quicksheet-<name>-ext`.
