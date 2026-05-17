# Show HN draft — QuickSheet

Drafted 2026-05-14, revised 2026-05-17 twice — once for the differentiator line per `research/adjacent-pitch-teardown.md`, and once for the iteration-speed paragraph per `research/loved-features-inverse-teardown.md`. Author posts manually. Pick one title; tweak copy as you like.

---

## Title options (HN allows ~80 chars; "Show HN:" prefix recommended)

1. **Show HN: QuickSheet – my desktop wallpaper is a spreadsheet**
2. **Show HN: A terminal spreadsheet that doubles as your desktop wallpaper**
3. **Show HN: QuickSheet – transparent interactive grid replaces my wallpaper**
4. **Show HN: I replaced my wallpaper with a spreadsheet (C#, zero NuGet deps)**

Recommended: **#1**. Short, concrete, slightly weird — the wallpaper angle is the hook.

## URL

`https://github.com/cemheren/QuickSheet`

## First comment (post as soon as the submission goes live — this is the pitch)

> This rewrite leads with personal-itch framing per the research in `research/show-hn-tui-patterns.md` — the Bagels (Show HN: TUI expense tracker, 283 pts) opening pattern. Two sentences of "I built this for myself because X" before any feature list.

```
I'm Akif. My wallpaper has been a static image I never interact with for years, and at the same time I always had a small handful of things I wanted at hand — notes, app launchers, a couple of URLs I open every morning, small running totals. I wanted those things to *be* the wallpaper, so I built QuickSheet for myself and it's been the surface I use most.

The simplest way to say what it is, vs Rainmeter / Conky / Übersicht / GeekTool: **the data is a CSV file. Cells can run shell commands. Same file on Windows or Linux.**

Two run modes share one CSV file: a normal terminal TUI in any shell, and a `--desktop` mode that embeds the same grid as the wallpaper (Win32 WorkerW on Windows, `_NET_WM_WINDOW_TYPE_DESKTOP` on X11).

Cell prefixes are the whole feature set:
- `r: code .` — runnable command. Press Enter to launch.
- `i: ping example.com` — subprocess output streams into the cell live.
- `s: 4,7,9,3,8,12` — renders as ▂▃▆▁▅█ sparkline. Also `s: A1::A10` for a range.
- `L: <cell>, 5m` — loops a target cell on an interval.
- `ext: github:user/repo` — clones the repo and registers a new prefix at runtime.
- Any URL — highlighted, opens on Enter.

What I open less since this exists: Trello (replaced by a column of `r: code .` cells for the morning repos), htop (`sysmon:` ext pinned in a corner), browser tabs for service status pages (`apistatus:`, `tls:`, `health:`). The wallpaper sits *behind* everything and is glanceable on alt-tab — no app to launch, no tab to find. The thing that surprised me using my own tool: I never *open* QuickSheet, it's always there.

A few constraints I held to and am glad I did:
- Zero NuGet dependencies. All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. Clone, build, run.
- CSV is the only persistence format. Open it in Excel or vim, same data.
- Cross-platform via OS-conditional `#if`s in shared files and per-OS folders excluded by the csproj, not an abstraction layer.

Honest disclosures: side project, written largely with AI assist (the zero-deps rule kept the design honest because you can't paper over a bad idea with a package). Wayland is the open problem — XWayland passes the X11 hint through but compositors don't honor it; layer-shell is the path forward and I haven't shipped that yet (issue #3 on the repo).

Repo: https://github.com/cemheren/QuickSheet
60-second tour: https://github.com/cemheren/QuickSheet/blob/main/docs/tour.md
```

## Pre-written reply blocks (post when the obvious questions come up)

### "Why not VisiData / sc-im / NeoVim with a spreadsheet plugin?"

```
Honest answer: none of those embed as the desktop wallpaper, which is the actual differentiator. If all I wanted was a TUI spreadsheet I'd use sc-im — it's mature and the keybindings are great. QuickSheet's design pressure came from wanting the grid to be *visible while doing other things*, not visible *only while focused on it*. The cell-prefix extension system (any cell prefix can be backed by a separate git repo of any language that talks JSON-lines on stdin/stdout) is the second differentiator and falls out of the wallpaper use case — you want compact, glanceable widgets, not full apps.
```

### "Why .NET? Why not Rust / Go?"

```
Two specific reasons. (1) The Win32 WorkerW wallpaper trick needs a window with a real HWND that can be `SetParent`ed into another HWND. .NET's WinForms gives me that with a one-line `new Form()` and lets me write the Win32 P/Invoke directly in the same project. Rust/Go can do this but the path is rougher. (2) The Linux side wants raw X11 P/Invoke too — `libX11.so.6` for the window, `libXft.so.2` for fonts — and .NET's `DllImport` is the same shape on both OSes, so the cross-platform conditional compilation pattern stays clean. Once that was in place I never had a reason to leave .NET.
```

### "Zero NuGet deps is dogma, not a design principle."

```
It's a forcing function more than a principle. The point isn't "packages are bad," it's that *not having packages* keeps the surface area small enough to read in a sitting and supply-chain-trivial to clone-and-run. When I needed ConPTY I wrote ~100 lines of P/Invoke instead of a 700-package dependency tree. The lift was less than I expected. The cost of the rule is real (the X11 font rendering took longer than `Avalonia.Controls.TextBlock` would have), but it bought a project I'd be comfortable shipping a binary of to a stranger.
```

### "Wayland?"

```
That's the open issue (#3 on the repo). `_NET_WM_WINDOW_TYPE_DESKTOP` doesn't survive XWayland — most compositors don't honor it. The right path is `wlr-layer-shell` with the background layer on wlroots-based compositors (Sway, Hyprland). I haven't shipped that yet; if anyone here has driven layer-shell from a non-Wayland-native language, I want to hear what was painful.
```

## Tips for the actual submission

- Post Tuesday–Thursday, 7–10am Pacific. Best window for `/newest` visibility.
- Do NOT title-bait. HN punishes hype.
- Reply to every comment in the first 2 hours. Engagement helps ranking.
- If someone asks "why not [other tui spreadsheet]?", answer with what's *different*, not what's better. The wallpaper-embedding + extensions angle is the differentiator.
- Have one screenshot or GIF ready to link in a reply — wallpaper mode is the visceral hook. The repo README has stills; a 10-second GIF would be even better.

## Lobsters / Reddit reuse

Title and first comment above port directly to:
- Lobsters (use `programming`, `dotnet` tags)
- r/commandline (drop "Show HN:" prefix, keep body)
- r/dotnet (lean harder on the zero-NuGet / P/Invoke angle in the body)
- r/coolgithubprojects (one-line caption + screenshot is enough)
