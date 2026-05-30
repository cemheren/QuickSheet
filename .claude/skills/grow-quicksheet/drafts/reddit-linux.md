# Reddit r/linux draft — QuickSheet

Drafted 2026-05-30. Subreddit rules: r/linux (2M+ members) allows self-posts about open-source Linux software. Flair: "Software Release" or "Open Source" depending on available options. They reward genuine technical substance over marketing. Lead with the X11 implementation detail — this crowd cares about how it works under the hood.

## Title

**I replaced my X11 wallpaper with an interactive spreadsheet grid — it uses _NET_WM_WINDOW_TYPE_DESKTOP to sit behind every window**

(Alternative: `Open-source desktop wallpaper that's actually an interactive CSV grid — X11 + .NET 9, zero dependencies`)

## Body

```
I've been working on QuickSheet — a .NET 9 app that turns your desktop into a transparent, interactive spreadsheet grid. It uses raw X11 P/Invoke (libX11.so.6, libXft.so.2) and sets `_NET_WM_WINDOW_TYPE_DESKTOP` on the window to make it behave like the wallpaper layer — sits behind every window, no taskbar entry, survives Win+D / show-desktop.

The grid is always there. Click the desktop → edit a cell. The file is plain CSV, autosaved every 5 seconds. Same data opens in LibreOffice, vim, csvkit, whatever.

**What makes it more than just "cells on a background":**

- `r: code .` — cell becomes a launcher. Enter fires the command.
- `i: sensors` — inline subprocess output, piped back into the cell live.
- `s: 4,7,2,9,3` → `▃▆▁█▂` — Unicode sparklines in a cell.
- `L: A5, 5m` — loop a cell every 5 minutes (great for `i: curl -s ...` status checks).
- `ext: github:user/repo` — extensions are separate processes speaking JSON-lines. 69+ available for weather, stocks, Docker, k8s pods, GitHub PRs, RSS, etc.
- Column Σ, row Π, multi-cell select + batch-execute.

**Technical details Linux users might care about:**

- Zero NuGet dependencies. All platform interop is hand-rolled P/Invoke — no native shim layers, no Electron, no GTK, no Qt.
- Renders with Xft (libXft.so.2) for antialiased text directly on the X11 surface.
- Keyboard input via XGrabKey / XNextEvent loop. Mouse via ButtonPress/MotionNotify.
- Single self-contained binary (`dotnet publish -r linux-x64`). ~30 MB, runs anywhere .NET 9 runs.
- Cross-platform: same codebase compiles for Windows (WorkerW embedding via WinForms) via conditional TFMs.

**Wayland caveat:** This uses X11 directly — it won't work on pure Wayland sessions. XWayland might work but isn't tested. The app detects `WAYLAND_DISPLAY` and prints a warning. If you're on Wayland, it falls back to a regular windowed mode. Wayland's layer-shell protocol could theoretically support this, but nobody's implemented it yet (PRs welcome — issue #3 tracks this).

**Install:**

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop
```

Requires .NET 9 SDK. Or grab the self-contained linux-x64 binary from the releases page — no SDK needed.

Add to your `~/.config/autostart/` for login persistence (XDG autostart .desktop file example in the docs).

Repo: https://github.com/cemheren/QuickSheet — MIT, zero deps, side project written with AI assist.

I'm curious what window managers people use that support `_NET_WM_WINDOW_TYPE_DESKTOP` well. It works on i3 (floating), KDE, XFCE, MATE. Haven't tested every DE — feedback welcome.
```

## Posting notes

- **Flair:** "Software Release" or equivalent. Check current r/linux flair options before posting.
- **Timing:** Weekday morning US-East (Mon–Wed, 8–10am ET) tends to get traction. Avoid weekends.
- **Engagement hook:** The Wayland question and the "which WMs support this?" question are genuine — they invite comments from people with niche setups.
- **Do NOT claim "production-ready" or "thousands of users."** The README already says "side project, AI-assisted." Honesty plays well on r/linux.
- **Screenshot:** Include one showing the grid behind terminal windows. The existing `desktop-wallpaper-commands.png` from the repo is suitable if it shows Linux.
- **If challenged on .NET on Linux:** Point to the zero-deps, single-binary story. The "but .NET is Microsoft" sentiment still exists on r/linux — counter with "it's MIT, it's a 30MB self-contained binary, there's no telemetry, no NuGet supply chain."
- **If asked about Conky comparison:** Conky renders pre-configured widgets (text/bars/graphs from system info). QuickSheet renders editable cells you type into. Different primitive — Conky is read-only; QuickSheet is interactive. Both can show system info, but only QuickSheet lets you *type something back*.

## Why r/linux specifically

- 2M+ members, large overlap with X11 power users who customize their desktop.
- The `_NET_WM_WINDOW_TYPE_DESKTOP` mechanism is well-known in this community (Conky uses it, desktop widget tools use it). Naming it in the title signals "this person knows what they're doing."
- r/unixporn is for screenshots of *riced* setups — r/linux is for discussing the software itself. This is a "here's the tool" post, not a rice post.
- Cross-pollinates well: r/linux readers who also browse r/i3wm, r/swaywm (the Wayland angle invites them), r/commandline.
