# Show HN: QuickSheet – Your desktop wallpaper is now an interactive spreadsheet

## Title options (pick one)

1. **Show HN: QuickSheet – Turn your desktop wallpaper into a live spreadsheet** (preferred)
2. Show HN: QuickSheet – A zero-dependency .NET spreadsheet that replaces your wallpaper
3. Show HN: I replaced my wallpaper with a transparent spreadsheet grid

## First comment (post immediately after submission)

Hi HN! I built QuickSheet because I wanted something always-visible on my desktop for quick notes, links, and app launchers — without opening yet another window.

**What it does:** Replaces your wallpaper with a transparent, interactive grid. You type directly into cells that sit behind your windows. Everything autosaves every 5 seconds.

**Key features:**
- `r: code .` in a cell → runnable command. Multi-select cells and hit Enter to batch-launch.
- Paste URLs → they become clickable hyperlinks.
- Auto-sum (Σ) per column, auto-product (Π) per row — no need to open a calculator.
- Extension system via JSON-lines protocol. There's a Copilot extension that brings AI into cells, and a weather widget.
- Works as both a desktop wallpaper app (Linux X11 + Windows) and a regular terminal TUI.

**Technical choices:**
- Zero NuGet dependencies. All native interop (X11 P/Invoke, WinForms, ConPTY) is hand-written. This was a deliberate supply-chain decision — you can audit everything.
- .NET 9, single-project build.
- CSV is the persistence format. No database, no JSON state files.
- Linux uses raw X11 with `_NET_WM_WINDOW_TYPE_DESKTOP` to sit below icons. Windows uses WorkerW embedding.

**What I use it for daily:**
- Morning routine: multi-select my project launchers, hit Enter, everything opens.
- Quick scratch notes during meetings (always visible, never buried).
- Hyperlink dashboard — my most-used internal URLs in a grid I can see anytime.
- Tracking small numbers (sprint points, daily counts) with auto-sum.

It's MIT licensed, ~4k LOC. Happy to answer questions about the X11/Win32 embedding tricks or the extension protocol.

https://github.com/cemheren/QuickSheet

## Timing notes

Best times to post on HN (US Pacific):
- Tuesday–Thursday, 8:00–9:00 AM PT (peak engagement)
- Avoid weekends and Monday mornings

## Pre-flight checklist before posting

- [ ] README has a compelling first paragraph ✓
- [ ] Quick Start works in ≤3 commands ✓
- [ ] At least one screenshot visible above the fold ✓
- [ ] GitHub topics set ✓
- [ ] Consider: record a 30-second GIF/video of desktop mode (high impact for HN)
- [ ] Consider: have a friend upvote within first 30 min (organic only)
