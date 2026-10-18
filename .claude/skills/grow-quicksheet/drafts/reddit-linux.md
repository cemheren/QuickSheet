# Reddit r/linux draft — QuickSheet

Drafted 2026-06-14. Subreddit rules: r/linux (3M+ members) accepts FOSS project posts
tagged as `[Project]`. Community rewards: technical depth, Linux-native approach, honest
scope. Dislikes: cross-post spam, "look at my side project" without substance, posts
that don't explain WHY it's Linux-specific.

**Distinct from other drafts:**
- `reddit-commandline.md` → terminal/TUI angle, cross-platform emphasis.
- `unixporn-rice.md` → screenshot-first, DE-tagged rice post.
- This → Linux integration angle: X11 P/Invoke, `_NET_WM_WINDOW_TYPE_DESKTOP`,
  why this isn't Conky. Technical readers, FOSS credibility.

**Position in launch sequence:** After r/unixporn (which needs the screenshot),
but can go before or after r/commandline depending on timing. Different audience
overlap — r/linux readers are more likely to ask about Wayland roadmap, packaging,
and distro support.

---

## Title

```
[Project] QuickSheet — interactive spreadsheet that embeds itself as the X11 desktop wallpaper (.NET 9, zero deps, MIT)
```

Alternative (if mods flag "Project" tag misuse):

```
QuickSheet: a spreadsheet that replaces your wallpaper on X11 — built with raw P/Invoke to libX11
```

---

## Body

```
I built a terminal spreadsheet that can embed itself as the Linux desktop wallpaper
using X11's `_NET_WM_WINDOW_TYPE_DESKTOP` hint. It sits behind all windows, receives
click events, and autosaves to CSV every 5 seconds.

**What it is:** A grid of cells on the actual desktop (not a separate app window you
minimize). You type notes, commands, URLs, formulas. It stays visible whenever you see
your wallpaper.

**How it works on Linux:**

The X11 embedding is done through raw P/Invoke to `libX11.so.6` and `libXft.so.2` —
no GTK, no Qt, no Electron. The window sets:

- `_NET_WM_WINDOW_TYPE_DESKTOP` → compositors place it below all windows
- `_NET_WM_STATE_BELOW` + `_NET_WM_STATE_SKIP_TASKBAR` → stays behind, no Alt+Tab
- Click-through is managed by selectively mapping/unmapping input on cells

Rendering is Xft text drawing directly to an X11 window. ~900 lines of C# P/Invoke
for the entire Linux platform layer.

**What cells can do:**

- `r: code .` → launches command on Enter (multi-select = batch launcher)
- `i: sensors` → live subprocess output piped back into the cell
- `s: 4,7,9,3,8,12` → sparkline: `▂▃▆▁▅█`
- `ext: github:user/repo` → extensions: JSON-lines over stdin/stdout
- Column sums (Σ), row products (Π), cell-range references
- Any URL → clickable hyperlink

**Hard constraints I'm holding:**

- Zero NuGet dependencies. All native interop is hand-written.
- MIT license. Single repo clone → `dotnet build` → run.
- CSV is the only persistence format. Your data is a text file.
- Cross-platform: same codebase compiles for Windows (WinForms/WorkerW embedding)
  and Linux (X11) via conditional TFMs.
- 120+ extensions across two orgs, all also zero-dep.

**What it is NOT:**

- Not Conky (Conky renders static text from Lua scripts; this is an interactive
  grid you type into).
- Not a widget layer (no drag-and-drop widgets; it's a spreadsheet with cell
  prefixes that trigger behavior).
- Not Wayland-compatible yet (issue #3 tracks this; the EWMH hints don't exist
  in Wayland's security model, so it needs a different approach — likely
  layer-shell or portal-based).

**Requirements:** .NET 9 SDK, X11 session. Works on GNOME/Xorg, KDE/X11, Xfce,
i3, bspwm, etc. Wayland compositors won't embed the window — it falls back to
a regular window on top.

Repo: https://github.com/cemheren/QuickSheet

Feedback welcome — especially from anyone who's worked with layer-shell protocols
on Wayland or has opinions on whether the `_NET_WM_WINDOW_TYPE_DESKTOP` approach
is the right one vs. using `_NET_WM_STRUT` or a background-drawing protocol.
```

---

## First comment (post within 5 minutes)

```
A few things I'd add that didn't fit the post:

1. Extensions are separate repos you install with `ext: github:user/repo`.
   QuickSheet clones the repo, builds/runs the binary, and communicates over
   JSON-lines stdin/stdout. There are 120+ already covering weather, stocks,
   RSS, Docker status, system monitoring, GitHub notifications, etc.

2. Theming: Ctrl+T cycles through 10 built-in themes (Nord, Dracula, Gruvbox,
   Synthwave, etc.). Themes are just named-color palettes applied to the X11
   draw calls.

3. Autosave is every 5 seconds to a CSV file. You can point it at any path:
   `dotnet run -- myfile.csv --desktop`. Same CSV opens in LibreOffice or vim.

4. The whole Linux platform layer is ~900 LOC of P/Invoke:
   https://github.com/cemheren/QuickSheet/tree/main/Platform/Linux

Happy to answer questions about the X11 approach or the extension protocol.
```

---

## Anticipated questions & replies

**Q: "Why .NET on Linux? Why not C/Rust/Go?"**

> Fair question. The project started on Windows with WinForms. When I added Linux
> support, the P/Invoke model meant I could call libX11 directly without porting
> the entire app. .NET 9 on Linux is first-class — single `dotnet build`, no
> Mono, AOT-ready. The zero-deps rule means there's no NuGet supply chain to
> worry about — it's just the SDK + system libs.

**Q: "Wayland when?"**

> Tracked in issue #3. The honest answer: `_NET_WM_WINDOW_TYPE_DESKTOP` has no
> equivalent in Wayland's security model. Options being explored:
> - wlr-layer-shell (sway, Hyprland) — works for wlroots compositors only
> - xdg-desktop-portal + background service
> - Compositor-specific plugins
>
> Contributions welcome. The abstraction layer (`IDesktopHost`) makes it clean
> to add a third platform impl.

**Q: "How is this different from Conky?"**

> Conky renders read-only text/graphics from Lua config scripts. QuickSheet is
> interactive — you click into cells, type, edit, run commands, install extensions.
> Think of it as "Excel meets Conky" rather than "Conky with a grid."

**Q: "Does it work with a compositor / transparency?"**

> The X11 window has no true alpha channel (Xft draws opaque text on an opaque
> background). Compositors like Picom can apply per-window opacity rules via
> `opacity-rule` in the config. The grid renders with a solid background color
> from the selected theme.

---

## Timing notes

- Best posted Tuesday–Thursday, 14:00–17:00 UTC (peak r/linux activity).
- Avoid weekends (lower engagement, posts get buried by memes in adjacent subs).
- Do NOT cross-post to r/commandline or r/unixporn on the same day. Space 3+ days.

## Success metrics

- 50+ upvotes = solid reception for a project post on r/linux.
- 100+ upvotes = strong; will hit front page of r/linux for ~12 hours.
- Key signal: comment quality > upvote count. Technical questions about X11/Wayland
  or "I'd use this if..." comments are more valuable than karma.
