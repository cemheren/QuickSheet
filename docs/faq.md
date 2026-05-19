# QuickSheet FAQ

Common questions newcomers ask. If you have one that isn't here, [open an issue](https://github.com/cemheren/QuickSheet/issues).

## Why not just use Excel / Google Sheets?

QuickSheet isn't trying to be Excel. It's a small grid that lives on your **desktop wallpaper**, behind your windows, with cells that can run commands, stream subprocess output, open URLs, and host extensions. Excel doesn't do any of that. A typical QuickSheet has a few rows of `r:` launchers, a few `i:` live-output cells, and some columns of numbers. If you need pivot tables, charts, or 10,000 rows, use Excel.

## How is this different from VisiData / Modin / sc-im?

VisiData and sc-im are excellent **data-analysis** TUIs — you load a CSV and slice it. QuickSheet is a **wallpaper / dashboard surface** — the grid is always there, between sessions, with live cells. Different use case. (And there's nothing stopping you from running VisiData on a CSV that QuickSheet writes.)

## Why .NET?

The project started in C# because the author was using Windows daily and wanted WorkerW desktop embedding without a heavy GUI framework. .NET 9 ships a small AOT-friendly runtime that handles cross-platform builds via OS-conditional TFMs in the csproj. The hard rule is **zero NuGet dependencies** — all native interop (X11, WinForms, ConPTY) is hand-written P/Invoke, so the dependency surface is just the .NET SDK itself.

## Will this run on macOS?

No native macOS port yet. QuickSheet is wallpaper-only, and the per-platform hosts are Windows (`WorkerW`) and Linux (raw X11). Wiring up an AppKit host would be a sizable but contained piece of work — contributions welcome.

## Wayland?

Not supported. The wallpaper integration requires the X11 `_NET_WM_WINDOW_TYPE_DESKTOP` hint, which Wayland compositors don't implement. The program prints a warning when it detects Wayland. Track [issue #3](https://github.com/cemheren/QuickSheet/issues/3) for the investigation.

## Where is my data stored?

A plain CSV file. Autosave path is `%APPDATA%/ExcelConsole/autosave.csv` on Windows, `$XDG_CONFIG_HOME/ExcelConsole/autosave.csv` (or `~/.config/ExcelConsole/autosave.csv`) on Linux. You can open it in any editor; that's the entire state.

## How do extensions get installed?

`ext: github:user/repo` in a cell clones the repo under `<config>/extensions/<owner>-<repo>/`, reads its `quicksheet-extension.json` manifest, and launches the entry command as a subprocess. The subprocess speaks JSON-lines on stdin/stdout. Full protocol in [docs/extensions.md](extensions.md).

## Do extensions get my credentials?

Each extension is a separate process — it has whatever credentials the **user** running QuickSheet has. The `gh`-based extensions (`ghpr`, `gitst`) use your existing `gh auth login`. The `docker:` extension talks to the local Docker socket if your user has access. Nothing is sent anywhere by QuickSheet itself; review an extension's source before installing.

## How do I write my own extension?

Any language with stdin/stdout works. Minimum: a register message on startup, a write message in response to each activate. The reference extensions are all under 300 LOC. See [docs/extensions.md](extensions.md) for the protocol and the "Build your own" section.

## Why so many extensions?

The extension network is the discoverability story — each extension is a tiny separate repo, and they fan out across the user's tax / dev / ops / personal verticals. The core stays small (no plugin loader, just a JSON-lines subprocess), and the network grows by adding repos rather than touching QuickSheet itself.

## Is there an installer?

Not yet. Today: `git clone` + `dotnet run`. A self-contained release build would be nice; if you've done it for a .NET 9 project before, a PR would land fast.

## Where do I report bugs or request features?

[GitHub issues](https://github.com/cemheren/QuickSheet/issues). For extensions, file in the extension's own repo (linked from [docs/extensions.md](extensions.md)).
