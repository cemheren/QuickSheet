# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

No tests, no linters. .NET 9 SDK required.

```bash
dotnet build ExcelConsole.csproj                                  # Build
dotnet run --project ExcelConsole.csproj                          # Run as desktop wallpaper
dotnet run --project ExcelConsole.csproj -- data.csv              # With a specific CSV
dotnet run -c Release --project ExcelConsole.csproj               # Release (recommended)
dotnet run --project ExcelConsole.csproj -- data.csv --export-md data.md  # Headless: CSV → Markdown
./scripts/build-libghostty.sh   # optional (dotnet build downloads a prebuilt copy): build libghostty-vt (Zig 0.17) into native/ (Linux; .ps1 on Windows)
```

`Program.cs` always launches the platform `IDesktopHost` (after handling headless `--export-*` flags). There is no TUI mode — the project is desktop-wallpaper only.

## Architecture

- **`GridManager`** — pure data: cell grid, CSV I/O, column sums (Σ), row products (Π). No platform/UI deps. `SaveToCsv` preserves rows/cols beyond visible grid.
- **`Platform/IDesktopHost`** — `Run(csvPath)` + `Dispose()`. Per-platform impls embed grid as wallpaper. Autosaves every 5s.
- **`Platform/Windows/`** — WinForms host. `DesktopFormBase` does Z-order locking, Alt+Tab hide, Win+D detection. `NativeMethods.cs` Win32 P/Invoke. `WorkerW` embedding.
- **`Platform/Linux/`** — raw X11 P/Invoke (`libX11.so.6`, `libXft.so.2`). Sets `_NET_WM_WINDOW_TYPE_DESKTOP`. Requires X11 (Wayland warning emitted).
- **`Features/IMode.cs`** — modal input interface (Enter/Exit/Commit/HandleKeyEvent).
- **`InlineProcessManager`** — live subprocesses for `i:` cells. ConPTY on Windows, openpty + posix_spawn on Linux. PTY bytes are fed into an `IVtScreen`. Thread-safe (UI reads, bg threads write). Scrollback capped at ~5,000 lines (page-granular in libghostty).
- **`Terminal/`** — `IVtScreen` screen model for PTY output. `GhosttyVtScreen` wraps libghostty-vt (hand-written P/Invoke in `GhosttyNative`, loaded from `QUICKSHEET_LIBGHOSTTY`, next to the exe, or system path); `LegacyVtScreen` is the regex fallback when the lib is absent. `VtScreen.Create` picks between them.
- **`CellPrefix`** — parses `i: ` (inline output), `r: ` (runnable), `s: ` (sparkline render), `http(s)://` (hyperlink), and `{A1::C10}` cell-range refs.

## Cross-platform build

`.csproj` uses OS-conditional TFMs + defines:
- Windows → `net9.0-windows`, `UseWindowsForms`, define `PLATFORM_WINDOWS`
- Linux → `net9.0`, define `PLATFORM_LINUX`
- `Platform/Windows/**` and `Platform/Linux/**` are `<Compile Remove>`'d on the other OS.

In shared files use `#if PLATFORM_WINDOWS` / `#elif PLATFORM_LINUX` (see `Program.cs`, `InlineProcessManager.cs`).

## Conventions

- **Zero NuGet packages; one native dependency.** libghostty-vt is the only external dependency (see `libghostty.props`, `.github/workflows/libghostty.yml`). Do not add packages or other dependencies.
- **CSV is the persistence format.** No DB, no JSON state.
- **Cell prefixes drive behavior.** `r: cmd` runnable, `i: cmd` inline-process output, URLs auto-detected.
- **Namespaces.** `ExcelConsole` root; `ExcelConsole.Platform.{Windows,Linux}` for platform code; `ExcelConsole.Features` for mode interfaces; `ExcelConsole.Terminal` for the VT screen models.
- **libghostty-vt delivery.** CI (`.github/workflows/libghostty.yml`) builds it from Ghostty's pinned, SHA-256-verified source and publishes GitHub Release binaries. The csproj downloads the right one and verifies it against hashes in `libghostty.props` (skip with `-p:QuickSheetLibGhosttyDownload=false`). A library in `native/<rid>/` overrides the download. Never commit native binaries, and the app must still run without it (`LegacyVtScreen` fallback).
