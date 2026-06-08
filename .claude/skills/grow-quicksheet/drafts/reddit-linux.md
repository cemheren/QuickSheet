# r/linux submission draft — QuickSheet

Drafted 2026-05-30. Targets r/linux specifically (1.2M+ members). This audience cares about open-source philosophy, technical implementation details, and Linux-native tooling. Different tone from r/commandline (which is use-case focused) or r/unixporn (which is screenshot focused).

## Flair

`Open Source Project` (if available) or `Software/Applications`.

## Title

Options (pick one):

1. `I replaced my X11 wallpaper with an interactive spreadsheet — zero dependencies, all P/Invoke`
2. `QuickSheet: a .NET 9 spreadsheet that embeds behind desktop icons via _NET_WM_WINDOW_TYPE_DESKTOP`
3. `Side project: interactive CSV grid as your Linux desktop wallpaper (X11, raw P/Invoke, no Electron)`

Recommended: **#1** — leads with the "what" not the project name, includes the technical hook (P/Invoke) that r/linux respects.

## Body text

```markdown
I've been working on a side project called [QuickSheet](https://github.com/cemheren/QuickSheet) — it replaces your wallpaper with a transparent interactive grid. Click anywhere on the desktop to type a note, prefix a cell with `r: code .` to make it a launcher, paste URLs as clickable links, or run `i: htop` to embed a live subprocess in a cell.

**How it works on Linux:**

The X11 implementation uses raw P/Invoke to `libX11.so.6` and `libXft.so.2`. It creates a window and sets `_NET_WM_WINDOW_TYPE_DESKTOP` so your WM places it at the desktop layer — below all windows, above the actual wallpaper. No Electron, no GTK, no Qt. The entire binary is a single .NET 9 process with zero NuGet dependencies.

Rendering is direct Xft text drawing with font fallback. Input is grabbed from X11 key/button events. Autosaves to CSV every 5 seconds.

**What cells can do:**

- Plain text (notes, todos)
- `r: <command>` — runs on Enter (app launcher)
- `i: <command>` — embeds live process output in the cell (like `i: sensors` or `i: docker ps`)
- `s: A1:A10` — sparkline rendered from a column of numbers
- URLs auto-detected as clickable hyperlinks
- Column sums (Σ) and row products (Π)
- 69+ extensions installable by git URL (`ext: github:user/repo`) — weather, stocks, system monitoring, Docker status, GitHub Actions, etc.

**Limitations I'm upfront about:**

- X11 only. No Wayland support yet (issue #3 is open — `wlr-layer-shell` is the likely path but I haven't gotten there).
- .NET 9 SDK required to build. No distro packages yet (AUR/Nix expressions welcome as PRs).
- This is a side project with AI-assisted code. It works for my daily use but isn't battle-tested across every WM.

**Build & run:**

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop
```

Data is just a CSV file. Same file works on Windows too (Windows version uses WorkerW embedding behind desktop icons via WinForms + Win32 interop).

MIT licensed, happy to answer questions about the X11 implementation or the extension protocol.
```

## Comment strategy

- If asked about Wayland: "Issue #3 is tracking it. The window-type hint approach won't work — it'll need wlr-layer-shell or the ext-session-lock protocol. PRs welcome if someone knows that space better than I do."
- If asked about performance: ".NET 9 NativeAOT would be the next step for startup time. Right now it's ~200ms cold start on my machine (Ryzen 5, NVMe). Memory sits around 40-60MB."
- If asked about alternatives: "Conky is the closest comparison — but Conky is display-only (no interaction, no typing into it). This is editable. The trade-off is it's much newer and less polished."
- If asked about the .NET choice: "Started as a Windows project (WinForms embedding). When I ported to Linux, P/Invoke to libX11 turned out to be surprisingly clean from C#. The unsafe pointer math is gnarly but contained to ~2 files."

## Timing

- Post Tuesday or Wednesday morning UTC. r/linux weekday traffic is higher than weekends (opposite of r/unixporn).
- Avoid: kernel release days, major distro release days, Wayland controversy threads (they attract negativity and your X11-only project will catch stray fire).

## Why r/linux specifically

- The "zero dependency, raw P/Invoke to X11" angle is catnip for this audience. They respect projects that don't reach for Electron or heavy frameworks.
- The open-source philosophy (MIT, no telemetry, CSV-as-storage, git-URL extensions) aligns with community values.
- Wayland limitation is a risk — but being honest about it upfront ("issue #3 is open, PRs welcome") earns trust rather than getting attacked.
- The audience skews toward people who actually use their Linux desktops daily — exactly the people who'd benefit from an interactive wallpaper.

## Pitfalls to avoid

- Don't post as a link-only submission. r/linux prefers self-posts with context for projects.
- Don't lead with "I built this" — lead with what it does and how it works technically.
- Don't mention Windows first. Linux audience wants to know this is Linux-native, not a Windows app ported over.
- Don't oversell. "Side project", "works for my use", "PRs welcome" — humble tone lands better here than "production-grade."
- Don't compare to commercial tools. Compare to Conky (open-source, same niche) if anything.

## Expected outcome

r/linux is hit-or-miss for project posts. A well-framed technical post with honest limitations can hit 200-500 upvotes. The X11 P/Invoke angle is novel enough to drive discussion. Star conversion is lower than r/unixporn (which is screenshot-driven) but higher quality — r/linux users who star tend to actually try the project.
