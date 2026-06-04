# Console.dev Submission Draft

Submit at: https://console.dev/submit

---

## Tool name

QuickSheet

## URL

https://github.com/cemheren/QuickSheet

## Short description (for the form)

A zero-dependency .NET 9 app that replaces your desktop wallpaper with an
interactive spreadsheet grid. Cells run shell commands, display live subprocess
output, fetch data from 69+ extensions (weather, stocks, RSS, system stats),
and auto-save to a plain CSV. Works on Windows and Linux (X11). Think of it as
a persistent scratchpad/launcher/dashboard that lives behind your windows
instead of a static wallpaper.

## One-liner pitch

Your desktop wallpaper is a spreadsheet — notes, app launchers, live data, all
in a transparent grid with no dependencies.

## Why it's interesting (expanded — for the editorial review)

- **Novel UX concept.** Not another TUI that replaces your terminal — it
  replaces your *wallpaper*. The grid is always behind every window, always
  one click away.
- **Shell-native.** `r: code .` runs VS Code. `i: tail -f /var/log/syslog`
  streams live output into a cell. URLs become clickable hyperlinks.
- **Extension ecosystem.** 69+ community extensions (GitHub Actions status,
  weather, crypto prices, Docker stats, RSS feeds, dice roller) all via a
  simple JSON-lines stdio protocol. Install with one cell:
  `ext: github:cemheren/quicksheet-weather-ext`.
- **Zero dependencies, zero NuGet packages.** Clone → `dotnet build` → run.
  All platform interop is hand-written P/Invoke. Entire supply chain is the
  .NET SDK.
- **Cross-platform.** Windows (WinForms + WorkerW embedding) and Linux (raw
  X11 via P/Invoke with `_NET_WM_WINDOW_TYPE_DESKTOP`).
- **Data is just CSV.** Your spreadsheet is a CSV file you can version, grep,
  pipe through awk, or open in Excel.
- **Headless pipeline mode.** `--export-md`, `--export-json`, `--filter`,
  `--sort`, `--select`, `--transpose` — use it as a CSV swiss-army knife in
  shell scripts.

## Category suggestions

Developer Tools, CLI/TUI, Productivity, Desktop

## Tech stack

.NET 9, C#, P/Invoke (Win32 + X11), zero third-party dependencies

## Install

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop
```

Pre-built binaries (win-x64, linux-x64) available on GitHub Releases.

---

## Notes for submitter (user)

- Console.dev form is at https://console.dev/submit — just needs tool name,
  URL, and a short description. The expanded section above is for the
  editorial team if they reach out for more detail.
- No screenshot/image required by their form, but linking the repo screenshot
  in the description helps.
- Best time to submit: right after a release with binary assets attached (more
  polished first impression).
