# Agent 5 — apply reviewer findings (Sonnet)

Exact prompt (`$SCRATCH` = session scratch dir).

````text
You are fixing code-review findings in /home/user/QuickSheet (branch ccr-7ea14cd6-efxsg4; read /home/user/QuickSheet/CLAUDE.md first; zero NuGet packages, hand-written P/Invoke, no unsafe code). The last commit (`git -C /home/user/QuickSheet show HEAD`) migrated inline-process console emulation to libghostty-vt (Terminal/*.cs, ConPtyProcess.cs, Platform/Linux/PtyProcess.cs, InlineProcessManager.cs, csproj, scripts, docs). Do NOT commit. Make exactly the fixes below, nothing else. libghostty-vt 1.0.0 headers for reference: $SCRATCH/vt100/include/ghostty/vt/ . .NET SDK: $SCRATCH/tools/dotnet/dotnet .

## 1. BLOCKER — Windows build breaks without the native lib (ExcelConsole.csproj)
The Windows `<None Include="native/win-$(QsNativeArch)/ghostty-vt.dll" .../>` is a literal path, so MSBuild fails with MSB3030 when the file doesn't exist. Add `Condition="Exists('native/win-$(QsNativeArch)/ghostty-vt.dll')"` on that item (keep the Linux wildcard as is).
Also (finding 7): derive the arch from the target, not the host: if `$(RuntimeIdentifier)` is set, take the part after the last '-' (`x64`/`arm64`/`x86`); otherwise fall back to host OSArchitecture mapped X64→x64, Arm64→arm64, X86→x86 (lowercase anything else). Verify with a scratch csproj OUTSIDE the repo (copy the relevant ItemGroups; you can't build Windows here, but you can force the condition branch with a property) that a missing DLL no longer errors, and `dotnet msbuild ExcelConsole.csproj -getProperty:QsNativeArch -p:RuntimeIdentifier=linux-arm64` prints arm64.

## 2. MAJOR — PTY + emulator resized back and forth every frame
Callers request the same process with different sizes each frame: span pre-scan (Platform/Linux/DesktopWindow.cs:612, Platform/Windows/DesktopForm.cs:505) passes the span size, while the row loop (DesktopWindow.cs:679, DesktopForm.cs:569), the status bar (DesktopWindow.cs:904, DesktopForm.cs:774) and copy (DesktopForm.cs:1117) call `ResolveDynamicValue(r, c)` with the 120x30 defaults; `InlineProcessManager.EnsureRunning` now resizes whenever the size differs → two SIGWINCH/ResizePseudoConsole + two reflows per frame, and dirty forever so the render loop never idles.
Fix:
- In both DesktopWindow.cs and DesktopForm.cs change the private wrapper to `ResolveDynamicValue(int row, int col, int? ptyCols = null, int? ptyRows = null)`. Size used to START a process stays as today: `commandCols = IsJson ? 4096 : (ptyCols ?? 120)`, rows `ptyRows ?? 30`. Pass `resize: ptyCols.HasValue && !IsJson(cell)` through to EnsureRunning.
- `InlineProcessManager.EnsureRunning(int pointerRow, int pointerCol, string command, int ptyCols = 120, int ptyRows = 30, bool resize = false)`: when the process already exists with the same command, call `existing.Resize(...)` ONLY if `resize` is true. Update its doc comment: only the span pre-scan (the authoritative visual size) may resize a running process.
- Leave the span pre-scan call sites passing their sizes (they're the only ones that should resize). Don't change any other rendering logic.

## 3. MAJOR — single-cell i: can show empty
The single-cell renderers show text up to the first '\n'; the old code dropped empty lines, but the ghostty formatter emits leading blank rows (Windows `ping`/`ipconfig` output starts with CRLF). In Terminal/GhosttyVtScreen.cs GetText, after removing '\r', also strip leading empty lines (TrimStart('\n')); keep TrimEnd. (LegacyVtScreen already drops empty lines.)

## 4. MAJOR — scrollback byte cap truncates wide (4096-col JSON) terminals
A ghostty cell is 8 bytes and rows are full width, so 64 MiB at 4096 cols holds only ~2000 rows < the 5000-line target. In GhosttyVtScreen compute the byte cap from the size: `bytes = (ulong)cols * 8 * 2 * (ulong)(maxScrollbackLines + rows)` (×2 slack for style/grapheme metadata), clamped to [8 MiB, 512 MiB]. Apply in the ctor and recompute+re-set in Resize when cols change (store maxScrollbackLines in a field). Add a short comment explaining.

## 5. MINOR — ConPTY resize option
Add constant `OptResizePullScrollback = 41` (bool*, 1 byte; verify in terminal.h). Add a `TerminalSetBool` helper in GhosttyNative like FormatterSetBool. In GhosttyVtScreen ctor, `#if PLATFORM_WINDOWS` set it to false (ConPTY keeps its own screen buffer without scrollback; header says set false for ConPTY). Treat a non-success result as non-fatal (ignore) since it's a tuning option.

## 6. MINOR — Windows DLL search hijack (Terminal/GhosttyNative.cs ProbeLoad)
On Windows, drop the bare-name `NativeLibrary.TryLoad("ghostty-vt.dll")` fallback (it searches PATH and CWD). Windows only loads from QUICKSHEET_LIBGHOSTTY (must be an absolute path — check Path.IsPathFullyQualified on all OSes and ignore otherwise) and AppContext.BaseDirectory. Keep the bare-name system probe on Linux/macOS.

## 8. MINOR — LegacyVtScreen unbounded partial line
Cap the in-progress line buffer `_current` (keep only the last 64 KiB of chars when it grows beyond that). Also when stripping the partial line for GetText, drop a trailing INCOMPLETE escape sequence (e.g. text ending in "\x1b", "\x1b[3", "\x1b]0;title" without BEL) so a sequence split across reads doesn't flash "[3" — e.g. a second regex `\x1b(?:\[[0-9;?]*|\][^\x07]*|\(?)$` applied before the main strip on the partial line only.

## 9/10. NITs in GhosttyVtScreen.GetText
- Decode straight from native memory: `Marshal.PtrToStringUTF8(ptr, checked((int)len))` instead of copying into a byte[].
- If format_alloc fails 3 times in a row, clear `_dirty` anyway (keep returning the cache) so the render loop can idle; reset the failure counter on success.

## 11. NITs
- GhosttyNative.Init: merge the two identical catch blocks into one `catch (Exception ex)`.
- Remove `VtScreen.UsingGhostty` if nothing uses it (grep).
- CLAUDE.md: the InlineProcessManager bullet "Output capped at 5,000 lines." → "Scrollback capped at ~5,000 lines (page-granular in libghostty)."
- README.md: the Windows command line should be `powershell -ExecutionPolicy Bypass -File scripts/build-libghostty.ps1` (default ExecutionPolicy blocks running it directly). Check the .ps1 header says the same.

## Verify
- `$DOTNET build /home/user/QuickSheet/ExcelConsole.csproj` → 0 warnings, 0 errors.
- Compile-check the Windows path with a scratch project OUTSIDE the repo (net9.0, DefineConstants PLATFORM_WINDOWS, EnableDefaultCompileItems false, Compile Include ConPtyProcess.cs, InlineProcessManager.cs, Terminal/*.cs, plus a stub for CellPrefix.IsCommand if needed). DesktopForm.cs can't be compiled on Linux (WinForms) — review your edit there extra carefully by reading it and making it mirror the Linux change exactly.
- Scratch test (outside repo) for LegacyVtScreen: write "ab\x1b[3" then "1mred\x1b[0m\n" → GetText after first write must be "ab" (no "[3"), after second "abred". And a 200 KiB line with no newline keeps GetText length ≤ 64 KiB.
- Never write test projects inside /home/user/QuickSheet; clean bin/obj there are gitignored and fine.

Report back: per-finding what you changed (file:line), verification output, anything you deviated on and why.
````
