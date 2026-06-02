# awesome-ricing submission — draft (2026-06-02)

> User submits the PR manually. Do not push to fosslife/awesome-ricing.

## Target

- **Repo:** [fosslife/awesome-ricing](https://github.com/fosslife/awesome-ricing) (4.3k ★)
- **Section:** `Background setting utilities and generators`
- **Last updated:** 2026-03-31
- **Style:** `- [Name](url) - One-line description. (language)`

## Why it fits

The "Background setting utilities and generators" section lists tools that manage
or generate desktop wallpapers/backgrounds (feh, Nitrogen, komorebi, pywal, etc.).
QuickSheet's `--desktop` mode embeds an interactive spreadsheet directly on the
desktop as a transparent window-type-desktop layer — it *is* the wallpaper. This
is a natural and novel fit: instead of rotating images, it makes the desktop
functional.

## Diff line to add

Place alphabetically after `quickwall` / before `setroot`:

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive spreadsheet that embeds as your desktop wallpaper. Cell prefixes for runnable commands, live subprocess output, sparklines, and 30+ extensions. Zero dependencies, CSV persistence. Windows (WorkerW) + Linux (X11). (.NET 9 / C#)
```

## PR title

```
Add QuickSheet (interactive spreadsheet as desktop wallpaper)
```

## PR body

```
Adds [QuickSheet](https://github.com/cemheren/QuickSheet) under **Background setting utilities and generators**.

Unlike traditional wallpaper setters that rotate images, QuickSheet replaces your
desktop background with a live, interactive spreadsheet. It embeds itself behind
desktop icons using platform-native techniques:

- **Linux**: raw X11 P/Invoke with `_NET_WM_WINDOW_TYPE_DESKTOP`
- **Windows**: WinForms WorkerW embedding (same trick as Wallpaper Engine)

Features relevant to ricing:
- Built-in theme presets (Nord, Dracula, Gruvbox, Catppuccin, etc.)
- Sparkline-in-cell rendering, bold cells, auto-fit column widths
- `i:` prefix for live subprocess output on your desktop (htop, weather, etc.)
- 30+ community extensions (stock tickers, system monitors, dice rollers)
- Zero NuGet dependencies — all native interop is hand-written P/Invoke
- MIT licensed, actively maintained, cross-platform

Screenshot-friendly: the grid is transparent and minimal, designed to look good
in r/unixporn posts.

Link: https://github.com/cemheren/QuickSheet
```

## Also considered

- **`avtzis/awesome-linux-ricing`** (1k ★, active) — similar fit under a
  "Wallpaper" or "Desktop" section. Could be a follow-up submission after
  this one lands.
- **`fosslife/awesome-ricing` → CLI Tools → Informative** — QuickSheet
  could also fit here (it's an "information dashboard for your terminal/desktop")
  but Background section is the tighter match since the primary mode IS the
  wallpaper.
