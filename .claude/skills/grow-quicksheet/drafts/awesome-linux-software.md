# Awesome-Linux-Software submission — draft (2026-05-17)

> User submits the PR manually. Do not push to luong-komorebi/Awesome-Linux-Software.

## Target

Repo: https://github.com/luong-komorebi/Awesome-Linux-Software
File: `README.md`
Section: **Office / Productivity** (or **Other** if Office is too narrow).

## One-line entry (matches existing list formatting)

Pick whichever the maintainer's current line style is. Two variants:

**Compact:**
```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) — Interactive desktop wallpaper that's a live CSV spreadsheet grid. Cells run shell commands, sparklines, color highlights, and 40+ extensions over a JSON-lines protocol. MIT, zero deps, X11. `OpenSource`
```

**Slightly longer with badges (some sections do this):**
```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) — Your desktop becomes a spreadsheet. Transparent interactive grid behind every window: notes, runnable cells (`r: code .`), live subprocess output (`i: top`), sparklines, hyperlinks, and an extension protocol (40+ extensions for weather, stocks, system monitoring, Kubernetes pods, etc.). .NET 9, MIT, zero NuGet dependencies, X11. `OpenSource`
```

## PR title

```
Add QuickSheet — interactive desktop wallpaper spreadsheet
```

## PR body

```markdown
QuickSheet is a .NET 9 open-source desktop app (MIT, zero NuGet deps) that replaces
your wallpaper with a transparent, interactive CSV grid. Cells can:

- Hold plain notes (autosaves every 5s)
- Run shell commands on Enter (`r: code .`)
- Stream live subprocess output (`i: top -b -n1`)
- Render unicode sparklines (`s: A1::A10`)
- Open URLs / hyperlinks
- Install extensions via `ext: github:user/repo` — 40+ extensions for weather,
  stocks, Kubernetes pods, Docker health, GitHub PR queues, RSS, currency, etc.

Runs on Linux via raw X11 P/Invoke (sets `_NET_WM_WINDOW_TYPE_DESKTOP`). Wayland
prints a warning. Also runs in a plain terminal TUI when launched without
`--desktop`.

Repo: https://github.com/cemheren/QuickSheet
License: MIT
Section suggested: **Office** (or **Other** if a better section fits).
```

## Notes for the submitter

- This list uses a strict `OpenSource` / `Freeware` tag at the end of each entry — keep it.
- The maintainer often asks for the project to be runnable on Linux out of the box. `dotnet run -c Release --project ExcelConsole.csproj -- --desktop` is the single command.
- If the maintainer pushes back on "this is also a TUI not a Linux desktop app," reply by linking the wallpaper screenshot in the README; the desktop mode is the lead.
- Pair with a comment that this is *not* a Wayland-native app (X11 only) — being honest about the limitation makes the review faster.

## Why this list (vs awesome-selfhosted)

- awesome-selfhosted is for *server* software (Plex, Nextcloud, Pi-hole). QuickSheet is a desktop app — wrong list, would be rejected.
- Awesome-Linux-Software is *the* canonical "open-source apps for Linux desktop" list, ~22k stars, active maintenance, productivity section already has many cross-category tools (file managers, note apps, etc.) where QuickSheet's wallpaper-as-spreadsheet angle is genuinely novel.
- Higher hit-prob than awesome-tuis because it doesn't force the TUI framing.

## After submission

- If accepted: a permanent inbound link from a 22k-star list. Worth maybe 10–30 stars of long-tail trickle plus discovery in search results.
- If rejected: ask the maintainer which section they'd prefer or accept silently and move on.
