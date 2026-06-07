# QuickSheet for artists & designers

If you freelance illustrations, do commission work, or manage a creative studio —
**your desktop wallpaper can be your project board.** Commission queue, color
palettes, payment status, posting schedule — all glanceable behind Photoshop,
Clip Studio, or Blender without alt-tabbing to a Notion page.

> Think "Trello board, but it's your wallpaper, it's a CSV you own, and it never
> phones home."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/artist-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. Reboot-proof — the grid comes back
exactly where you left it.

## What goes on the wallpaper

The starter sheet at `examples/artist-dashboard.csv` covers three zones:

| Zone              | What it shows                                     | How                         |
|-------------------|---------------------------------------------------|-----------------------------|
| **Commissions**   | Client, deadline, status, rate, payment flag      | Plain cells — edit directly |
| **Color palette** | Hex codes + named color swatches via `c:` prefix  | Built-in color cells        |
| **Socials**       | Handles, next-post date, follower count           | Plain cells                 |

## Commission tracker

The simplest use: a row per open commission.

```
Commission,Client,Deadline,Status,Rate,Paid
Character portrait,@stormwitch_art,2026-06-20,🟡 WIP,$350,⬜
Album cover,Neon Drift Records,2026-06-28,⬜ Sketch phase,$800,⬜
Twitch emotes x6,ttv/galaxybrain,2026-06-15,🟢 Delivered,$180,🟢
```

Column sums (Σ) auto-total your rates. Sort by deadline with **Ctrl+B**. Search
clients with **Ctrl+F**.

## Color palettes on your wallpaper

Use the built-in `c:<color>:` prefix to paint swatches directly in cells:

```
c:red: ████
c:cyan: ████
c:yellow: ████
c:magenta: ████
```

Keep hex codes in an adjacent column for copy-paste into your art tool. Your
palette reference is always visible — no switching to Coolors or Adobe Color.

## Posting schedule

Track your social media calendar without a separate app:

```
Platform,Handle,Next post,Content idea
Instagram,@artof_skyline,2026-06-08,WIP timelapse
Bluesky,@skyline.art,2026-06-09,Finished commission reveal
ArtStation,skyline_studio,Portfolio update,New cover art
```

## Extensions that fit the creative workflow

| Extension | What it does | Install |
|-----------|-------------|---------|
| [`rss`](https://github.com/cemheren/quicksheet-rss-ext) | Art blog / inspiration feed on your wallpaper | `ext: github:cemheren/quicksheet-rss-ext` |
| [`weather`](https://github.com/cemheren/quicksheet-weather-ext) | Plein-air painting conditions at a glance | `ext: github:cemheren/quicksheet-weather-ext` |
| [`tz`](https://github.com/cemheren/quicksheet-tz-ext) | Time zones for international clients | `ext: github:cemheren/quicksheet-tz-ext` |
| [`1099`](https://github.com/cemheren/quicksheet-1099-ext) | Quarterly tax estimate for freelance income | `ext: github:cemheren/quicksheet-1099-ext` |
| [`mileage`](https://github.com/cemheren/quicksheet-mileage-ext) | IRS mileage deduction for art fair travel | `ext: github:cemheren/quicksheet-mileage-ext` |

Install any extension by placing `ext: github:cemheren/<name>` in a cell.

## Why wallpaper > Notion / Trello

- **Always visible** — commission deadlines are behind Clip Studio, not buried
  in a browser tab you forgot about.
- **Zero cloud** — client names, rates, and contact info stay in a CSV on your
  machine. No third-party sync.
- **Instant** — no loading spinner, no login, no "free tier limit reached."
- **Hackable** — need a cell that checks your Etsy shop stats or Ko-fi balance?
  Write a 50-line extension in any language
  ([extension-protocol.md](extension-protocol.md)).

## Make it yours

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula,
Synthwave, Gruvbox, Monokai, HotdogStand. Synthwave and Dracula pair well with
dark art workflows; Light is clean for client calls.

## What's missing (be honest)

- No image thumbnails in cells — it's a text grid, not a canvas.
- No direct integration with art platforms (DeviantArt, ArtStation API). You can
  build a custom extension if you need it.
- No calendar sync — deadlines are manual. That's intentional: you own the data.

## See also

- [Full extension directory](extensions.md)
- [Keyboard shortcuts](keyboard-shortcuts.md)
- [Extension protocol spec](extension-protocol.md) — build your own in an afternoon
- [for-traders.md](for-traders.md) — if you also trade
- [for-students.md](for-students.md) — if you're in art school
