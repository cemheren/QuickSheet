# Phase 2 — libghostty-vt as QuickSheet's one dependency

## Decision (maintainer)
"Break the no-dependency rule and take a single dependency on ghostty; the build requirement (Zig) is too much for users."

## Research (orchestrator)
- Users install from source (`git clone` → `dotnet run`; autostart does `git pull` + build), so the library must arrive at build time.
- Ghostty publishes **no** prebuilt Windows/Linux libraries and no NuGet package — only source tarballs (+ xcframework/wasm).
- Third-party NuGet options inspected (contents, nuspec, MSBuild targets):
  - `RoyalApps.RoyalTerminal.GhosttySharp.Native.{Win64,Linux64}` (~20k downloads): stable 0.7.2 ships pre-1.0 `libghostty-vt.so.0` (our binding requires ABI 1.x); 1.x build only in `0.8.0-preview.1`.
  - `DeBlasis.GhosttyVt`: only daily `-ci` prereleases; no linux-arm64.
  - `LibGhostty.Net`: net10.0 only, low adoption.
- Options put to the maintainer: (A) own CI build + pinned download, (B) own CI build + commit binaries, (C) third-party NuGet natives. **Chosen: A**, and **keep the regex fallback**.

## Delegation
| # | Agent (model) | Scope |
|---|---|---|
| 6 | CI workflow (Sonnet) | `.github/workflows/libghostty.yml`, script `ZIG_TARGET`/`LIBGHOSTTY_OUT`, `scripts/libghostty-smoke/` |
| 7 | csproj download (Sonnet) | `ExcelConsole.csproj` target `QsFetchLibGhostty`, `libghostty.props` |
| 8 | Docs (Haiku) | README, CONTRIBUTING, CLAUDE.md, CHANGELOG, faq, tour, website, copilot-instructions |
| 9 | Reviewer (Opus) | uncommitted delivery change, **before** first push (push publishes an immutable release) |
| 10 | Fixes (Sonnet) | reviewer findings in workflow/csproj/scripts/smoke |
| 11 | Docs fixes (Haiku) | SECURITY.md, CONTRIBUTING bump steps, website, growth-skill drafts |

## Key review catch (agent 9)
libghostty 1.0.0 `src/build/Config.zig:137-141` forces `abi = .msvc` on Windows when no ABI is given; upstream CI uses `-Dtarget=x86_64-windows-gnu`. Publishing as-is would have frozen DLLs that likely import the VC++ runtime. Fixed (GNU ABI + an import-allowlist gate) before anything was published. Orchestrator verified in source.

## CI history
1. Run 42895728511 (`42eb18a`) — all 4 legs failed, **nothing published** (publish needs every leg):
   - Windows: `uucode_generate.exe: FileNotFound`. Root cause (read in Zig 0.17 `lib/compiler/Maker/Step/Run.zig:2866-2893`): relative cache paths are rewritten relative to the helper's cwd, and Windows CreateProcess resolves them against the parent's cwd. Fix: absolute `--cache-dir`.
   - Linux: libraries built and passed the allowlist (libc, libm, librt); the export check failed because `readelf` truncates long versioned names without `-W`.
2. Run 42898319491 (`62eb0bc`) — all green; published `libghostty-vt-1.0.0-1`.
3. Pins (`99c6cdd`) cross-checked between `SHA256SUMS` and GitHub asset digests.

## Local verification after publish
- Clean `dotnet build`: downloads + verifies linux-x64 asset, 0 warnings.
- Smoke test vs the downloaded binary: 11/11 PASS.
- Real PTY (`PtyProcess`, 40×6), ghostty vs legacy:

| Output | libghostty-vt | legacy |
|---|---|---|
| cursor-up + erase redraw | `step 3/3 done` | 3 stacked lines |
| clear-screen refresh | current frame only | stale frame + new |
| alt-screen TUI exit | `before` / `after` | `full-screen UIafter` |
| `\r` progress, Unicode+color | correct | correct |
