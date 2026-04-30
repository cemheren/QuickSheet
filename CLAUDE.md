# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

No tests, no linters. .NET 9 SDK required.

```bash
dotnet build ExcelConsole.csproj                                  # Build
dotnet run --project ExcelConsole.csproj                          # Console TUI mode
dotnet run --project ExcelConsole.csproj -- --desktop             # Desktop wallpaper mode
dotnet run --project ExcelConsole.csproj -- --desktop data.csv    # Desktop mode with CSV
dotnet run -c Release --project ExcelConsole.csproj -- --desktop  # Release (recommended for desktop)
```

`Program.cs` dispatches: `--desktop` → platform `IDesktopHost`; otherwise → `SpreadsheetApp` (TUI).

## Architecture

Two run modes share one data layer:

- **`GridManager`** — pure data: cell grid, CSV I/O, column sums (Σ), row products (Π). No platform/UI deps. `SaveToCsv` preserves rows/cols beyond visible grid.
- **`SpreadsheetApp`** — console TUI via `System.Console`. Handles input, render, search, autosave to `%APPDATA%/ExcelConsole/autosave.csv`.
- **`Platform/IDesktopHost`** — `Run(csvPath)` + `Dispose()`. Per-platform impls embed grid as wallpaper. Desktop mode autosaves every 5s.
- **`Platform/Windows/`** — WinForms host. `DesktopFormBase` does Z-order locking, Alt+Tab hide, Win+D detection. `NativeMethods.cs` Win32 P/Invoke. `WorkerW` embedding.
- **`Platform/Linux/`** — raw X11 P/Invoke (`libX11.so.6`, `libXft.so.2`). Sets `_NET_WM_WINDOW_TYPE_DESKTOP`. Requires X11 (Wayland warning emitted).
- **`Features/IMode.cs`** — modal input interface (Enter/Exit/Commit/HandleKeyEvent).
- **`InlineProcessManager`** — live subprocesses for `i:` cells. ConPTY on Windows, pipe redirect on Linux. Thread-safe (UI reads, bg threads write). Output capped at 200 lines.
- **`CellPrefix`** — parses `i: ` (inline output), `r: ` (runnable), `http(s)://` (hyperlink), and `{A1::C10}` cell-range refs.

## Cross-platform build

`.csproj` uses OS-conditional TFMs + defines:
- Windows → `net9.0-windows`, `UseWindowsForms`, define `PLATFORM_WINDOWS`
- Linux → `net9.0`, define `PLATFORM_LINUX`
- `Platform/Windows/**` and `Platform/Linux/**` are `<Compile Remove>`'d on the other OS.

In shared files use `#if PLATFORM_WINDOWS` / `#elif PLATFORM_LINUX` (see `Program.cs`, `InlineProcessManager.cs`).

## Conventions

- **Zero NuGet deps.** Hard policy — supply chain risk. All native interop is hand-written P/Invoke. Do not add packages.
- **CSV is the persistence format.** No DB, no JSON state.
- **Cell prefixes drive behavior.** `r: cmd` runnable, `i: cmd` inline-process output, URLs auto-detected.
- **Namespaces.** `ExcelConsole` root; `ExcelConsole.Platform.{Windows,Linux}` for platform code; `ExcelConsole.Features` for mode interfaces.
