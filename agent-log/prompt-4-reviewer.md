# Agent 4 — reviewer (Opus)

Exact prompt given by the orchestrator (`$SCRATCH` = session scratch dir, outside the repo).

````text
You are a senior code reviewer. Review the most recent commit in /home/user/QuickSheet (branch ccr-7ea14cd6-efxsg4; `git -C /home/user/QuickSheet show HEAD` / `git diff HEAD~1`). Read /home/user/QuickSheet/CLAUDE.md for project rules (zero NuGet deps, hand-written P/Invoke, CSV persistence, cross-platform #if). DO NOT edit any repo files and DO NOT commit — review only; report findings.

## What the change does
Migrates inline-process (`i:` cells) console emulation from a per-chunk UTF-8 decode + '\n' split + ANSI regex strip, to libghostty-vt (Ghostty's VT emulator C library, stable C ABI 1.0, release 1.0.0) via hand-written P/Invoke, with a regex fallback (LegacyVtScreen) when the native lib isn't present. Files: Terminal/*.cs (new), ConPtyProcess.cs (Windows), Platform/Linux/PtyProcess.cs, InlineProcessManager.cs, scripts/build-libghostty.{sh,ps1}, ExcelConsole.csproj, .gitignore, docs.

## Authoritative C ABI reference
The libghostty-vt 1.0.0 headers are checked out at $SCRATCH/vt100/include/ghostty/vt/ (terminal.h, formatter.h, types.h, allocator.h, build_info.h, io.h). The Zig implementation (newer main, for semantics only) is at $SCRATCH/ghostty-build/src/terminal/c/. VERIFY every P/Invoke signature in Terminal/GhosttyNative.cs against the headers: parameter types/widths (uint16_t, uint32_t, size_t, C bool = 1 byte, enums = int), out-params, pointer-vs-value semantics for ghostty_terminal_set / ghostty_formatter_set option values, enum/option numeric values, ghostty_free allocator pairing, calling convention. Do not run any build of the ghostty source tree.

## Focus areas (correctness first)
1. ABI correctness of every DllImport and option constant (as above). A wrong width or value is a crash/memory-corruption bug.
2. Thread safety & lifetime: PTY reader threads call Write while the UI thread calls GetText/HasNewOutput; Dispose ordering (formatter freed before terminal; screen disposed after reader join; Write-after-Dispose); use-after-free risk if Dispose's Join times out; races on _exitMarker/_markerPending in PtyProcess; HasNewOutput semantics vs the render loop in Platform/Linux/DesktopWindow.cs (~line 481) and Platform/Windows/DesktopForm.cs.
3. Library loading in GhosttyNative: SetDllImportResolver semantics (one per assembly — does it interfere with X11Methods' DllImport("libX11.so.6") etc.? returning IntPtr.Zero falls back to default probing), probe order, failure handling, version check.
4. Behavioral regressions vs the old code: e.g. old code dropped empty lines — new GhosttyVtScreen keeps blank lines from the screen; old Linux appended "[exited N]"; JSON cells use 4096 cols (check callers of GetOutput for JSON parsing in GridManager/CellPrefix and whether formatter output changes break it); InlineProcessManager now resizes on every EnsureRunning call with a different size — could this thrash (e.g. ptyCols computed differently between the span pre-scan and normal cell render, causing resize every frame → SIGWINCH storms / reflow cost)? Check ALL callers of ResolveDynamicValue / EnsureRunning in both DesktopWindow.cs and DesktopForm.cs with the cols/rows they pass. This one matters — please trace concretely.
5. Performance: GetText formats whole scrollback (up to 5000 lines) when dirty — called per render frame per inline cell; is caching correct? Any per-frame allocation in hot paths?
6. LegacyVtScreen correctness (CR/LF/progress semantics, decoder flush, cap).
7. Build scripts: bash/PowerShell correctness, checksum verification really gates the build, temp cleanup, the zig flags (`--system`, `-Demit-lib-vt`, `-Dcpu=baseline`, `-Dstrip=true`) — check against PACKAGING-LIBGHOSTTY.md at $SCRATCH/ghostty-build/PACKAGING-LIBGHOSTTY.md, and the Windows output location (bin/ghostty-vt.dll).
8. csproj: the native copy ItemGroups, the Compile/None Remove of native/** and scripts/**.
9. Docs accuracy.

A .NET 9 SDK is at $SCRATCH/tools/dotnet/dotnet if you want to build (you may build the repo and scratch projects OUTSIDE the repo; never write inside /home/user/QuickSheet — use `-o` / obj paths outside or clean up obj/bin, which are gitignored anyway).

## Output format
Return a list of findings, most severe first. For each: severity (BLOCKER / MAJOR / MINOR / NIT), file:line, what's wrong, a concrete failure scenario, and a specific suggested fix. Only include findings you have verified by reading code/headers; mark anything uncertain as "PLAUSIBLE". Then a short list of things you checked and found correct. Be concise; no praise padding.
````
