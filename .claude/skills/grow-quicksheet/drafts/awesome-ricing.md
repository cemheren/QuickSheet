# Awesome-Ricing Submission Drafts — QuickSheet

Drafted 2026-06-15. Two ricing-focused lists where QuickSheet fits perfectly —
it literally replaces the desktop wallpaper with a live, interactive spreadsheet.

---

## 1. fosslife/awesome-ricing (4.3k ⭐)

**URL:** https://github.com/fosslife/awesome-ricing

**Section:** `Background setting utilities and generators`

QuickSheet embeds itself as the desktop wallpaper via X11 (`_NET_WM_WINDOW_TYPE_DESKTOP`)
on Linux and WorkerW embedding on Windows. It's not just a static image setter —
it's a live interactive grid you can type into, run commands from, and display
live subprocess output on. Perfect fit alongside `komorebi`, `xwinwrap`, and `wallset`.

**Diff line to add (alphabetical, after `quickwall`):**

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive terminal spreadsheet that embeds as a live desktop wallpaper. Runnable command cells, live subprocess output, sparklines, and an extension system. (C#)
```

**PR title:** `Add QuickSheet (live interactive wallpaper spreadsheet)`

**PR body:**

```
Adds [QuickSheet](https://github.com/cemheren/QuickSheet) under "Background setting utilities and generators".

QuickSheet is a .NET 9 terminal spreadsheet that can embed itself as a transparent,
interactive desktop wallpaper — you get a live grid on your desktop that you can
type into, run commands from, and display real-time subprocess output in.

Why it fits this section:
- Sets itself as the desktop background via X11 _NET_WM_WINDOW_TYPE_DESKTOP (Linux)
  or WorkerW embedding (Windows).
- Unlike static wallpaper setters, it's interactive — click cells, edit data,
  run shell commands, see live output.
- Supports theming (Nord, Dracula, Catppuccin, Tokyo Night) for rice-friendly aesthetics.
- Extension system lets you add custom cell prefixes (weather, stock tickers,
  system monitors, etc.) via JSON-lines protocol.
- MIT licensed, zero NuGet dependencies, cross-platform.

Repo: https://github.com/cemheren/QuickSheet
```

---

## 2. avtzis/awesome-linux-ricing (1.1k ⭐)

**URL:** https://github.com/avtzis/awesome-linux-ricing

**Section:** `Wallpapers > Utilities`

This list separates wallpaper utilities by compositor support (Wayland/X11 tags).
QuickSheet is X11-only, so it gets the `<sup>X11</sup>` tag.

**Diff line to add (alphabetical, after `hyprpaper` or in a new subsection):**

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet)<sup>X11</sup> - Live interactive spreadsheet as your desktop wallpaper. Editable cells, runnable commands, live subprocess output, sparklines.
```

**PR title:** `Add QuickSheet — interactive wallpaper spreadsheet`

**PR body:**

```
Adds [QuickSheet](https://github.com/cemheren/QuickSheet) to the Wallpapers > Utilities section.

QuickSheet is a .NET 9 spreadsheet that embeds as a transparent, interactive
desktop wallpaper on X11 (uses _NET_WM_WINDOW_TYPE_DESKTOP). It's not a
wallpaper image setter — it's a live grid you interact with directly on your
desktop background.

Features relevant to ricing:
- Theme presets: Nord, Dracula, Catppuccin Mocha, Tokyo Night.
- Transparent background blends with any setup.
- Sparkline cells, inline process output, hyperlinks.
- Extension system for custom widgets (weather, git stats, system monitors).
- Zero NuGet dependencies, raw X11 P/Invoke.

Currently X11-only on Linux (Wayland warning emitted if no X display).

MIT license. Screenshots in README.
```

---

## 3. Submission strategy

- **fosslife/awesome-ricing** is higher priority (4.3k stars = more eyeballs,
  and the "Background setting utilities" section is a direct match).
- **avtzis/awesome-linux-ricing** is secondary (1.1k stars, newer list, active).
- Both PRs should include a link to a screenshot showing QuickSheet as the
  desktop wallpaper — gated on the user having a wallpaper screenshot asset.
- If the list has a CONTRIBUTING.md, follow its PR template.
- One PR per list, submitted separately.

---

## Why this is high-ROI for QuickSheet

The ricing community (r/unixporn, awesome-ricing ecosystem) is the *exact*
audience for "spreadsheet as wallpaper." These users:
1. Actively customize their desktops.
2. Share screenshots that include background content.
3. Try new tools that make their setup look unique.
4. Overlap with terminal-tool enthusiasts who'd also use the TUI features.

Being listed in awesome-ricing puts QuickSheet in front of 4.3k+ stargazers
who are specifically browsing for wallpaper/background tools.
