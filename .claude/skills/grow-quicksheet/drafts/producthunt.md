# ProductHunt Launch Draft — QuickSheet

## Tagline (60 chars max)

> Your desktop wallpaper is now an interactive spreadsheet

## One-liner description

QuickSheet replaces your static wallpaper with a transparent, always-on grid — notes, app launchers, live data, and 69+ extensions, all without opening a window.

## Product description (longer)

Most developers have a monitor showing a wallpaper that does nothing. QuickSheet turns that dead space into a transparent, interactive grid that lives *behind* your windows.

**What you can do:**

- Click anywhere on the desktop to jot notes — autosaves every 5 seconds
- Prefix cells with `r: code .` to launch apps and scripts
- Paste URLs for a personal start page behind your windows
- Column sums, sparklines, inline subprocesses (`i: top -b -n1`)
- 69+ community extensions: weather, stocks, RSS, system monitoring, dice rolling, and more

**What makes it different:**

- Zero dependencies — the entire supply chain is the .NET SDK
- CSV is the data format — portable, human-readable, version-controllable
- Cross-platform: Windows (WinForms + WorkerW embedding) and Linux (raw X11 P/Invoke)
- Extensions are standalone executables that speak a JSON-lines protocol — write one in any language

**Who it's for:**

Developers, homelabbers, SREs, students, and anyone who lives in a terminal and wants their desktop to *do* something instead of just looking pretty.

## Topics / Categories

- Developer Tools
- Productivity
- Open Source
- Linux
- Windows

## Maker comment (first comment after launch)

Hey Product Hunt! 👋

I built QuickSheet because I was tired of my second monitor showing a static wallpaper while I worked. I wanted something lightweight and always-present — not another app window to manage.

The core idea: your desktop background is already the one thing that's *always there*. Why not make it interactive?

A few things I'm proud of:

1. **Zero NuGet dependencies.** All native interop (X11, Win32, ConPTY) is hand-written P/Invoke. The supply chain is just the .NET SDK.

2. **Extensions are any executable** that reads stdin and writes JSON-lines to stdout. Write one in Python, Go, Rust, whatever — no SDK required. The community has built 69+ so far.

3. **CSV as persistence.** Your data is a plain text file. `git diff` it, email it, pipe it through `awk`. No proprietary format.

The wallpaper embedding was the hard part — on Windows it uses the undocumented WorkerW window trick, on Linux it sets `_NET_WM_WINDOW_TYPE_DESKTOP` via raw X11. Both paths are ~200 lines of interop.

Would love feedback on the UX and extension protocol. What would you put on your desktop grid?

## Gallery images (descriptions for screenshots to capture)

1. **Hero shot** — Full desktop with QuickSheet visible: cells containing notes, a few `r:` launcher commands, some URLs highlighted, a sparkline, and normal windows overlapping parts of the grid.

2. **Extension ecosystem** — Grid showing live data: weather in one cell, stock ticker in another, system CPU/RAM sparkline, RSS headlines.

3. **Launch stack** — Close-up showing multi-selected `r:` cells about to launch a morning workflow (IDE, browser tabs, terminal sessions).

4. **Before/After** — Split image: left side is a boring static wallpaper, right side is the same desktop with QuickSheet active.

## Launch timing

- **Best day:** Tuesday or Wednesday (PH engagement peaks mid-week)
- **Best time:** 12:01 AM PST (PH resets at midnight Pacific)
- **Coordinate with:** HN "Show HN" post (same day, ~6:30–9:30 AM EST per cold-start research)

## Preparation checklist

- [ ] Demo GIF or video (15–30 seconds showing click-to-edit, launch, live data)
- [ ] 4 gallery images (see descriptions above)
- [ ] GitHub release with pre-built binaries (so non-devs can try it)
- [ ] Landing page or at minimum a polished README (current README is launch-ready)
- [ ] Hunter lined up (someone with PH followers to "hunt" the product)
- [ ] 5–10 supporters ready to upvote + leave genuine comments in first hour

## Post-launch actions

- Reply to every comment within 2 hours
- Share PH link on Twitter/X, Mastodon, LinkedIn
- Update repo description with "Featured on Product Hunt" badge if it gets featured
- Add PH badge to README if 100+ upvotes
