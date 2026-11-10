# Fix agent report

All requested fixes are applied and nothing is committed. The main Linux build is 0 warnings, 0 errors. The Windows-path compile check also builds clean. WinForms and the Windows MSBuild branch could not be run on Linux, so DesktopForm.cs and the Windows csproj item were checked by diff and by scratch tests only.

**1/7. csproj** (ExcelConsole.csproj, QsNativeArch PropertyGroup and the Windows `<None>` item)
- The Windows `<None>` item now has `Condition="Exists('native/win-$(QsNativeArch)/ghostty-vt.dll')"`. The Linux wildcard is unchanged.
- QsNativeArch now comes from the part of `$(RuntimeIdentifier)` after the last '-'. Without a RID it falls back to the host OSArchitecture lowercased, which gives x64, arm64 and x86.
- `dotnet msbuild ExcelConsole.csproj -getProperty:QsNativeArch -p:RuntimeIdentifier=linux-arm64` prints `arm64`. With no RID it prints `x64`.
- A scratch csproj outside the repo (forced `QsNativeArch=x64`, same Condition) builds with 0 errors and no MSB3030 when the DLL is missing. With a dummy DLL present, `ghostty-vt.dll` lands in the output folder.

**2. Resize every frame**
- `ResolveDynamicValue` is now `(int row, int col, int? ptyCols = null, int? ptyRows = null)` in Platform/Linux/DesktopWindow.cs (~1872) and Platform/Windows/DesktopForm.cs (~338).
- The start size is `IsJson ? 4096 : (ptyCols ?? 120)` for columns and `ptyRows ?? 30` for rows.
- `resize` is `ptyCols.HasValue && !isJson`, and it is passed to `EnsureRunning`.
- `EnsureRunning` in InlineProcessManager.cs (~94) takes a new `bool resize = false` parameter and calls `existing.Resize` only when it is true. The doc comment is updated.
- The span pre-scan call sites still pass their sizes and are the only ones that resize.
- The Windows edit is identical to the Linux one (compared by diff, not compiled).

**3.** Terminal/GhosttyVtScreen.cs GetText now does `Replace("\r","").TrimStart('\n').TrimEnd('\n',' ')`.

**4.** GhosttyVtScreen.cs has a new `ScrollbackBytes(cols, rows, maxLines)` helper using the formula you gave, clamped to [8 MiB, 512 MiB], with a comment. It is applied in the ctor, and `Resize` re-sets it when cols change. `_maxScrollbackLines` is stored in a field.

**5.**
- GhosttyNative.cs has `OptResizePullScrollback = 41` and a new `TerminalSetBool`. I checked 41 against terminal.h: `GHOSTTY_TERMINAL_OPT_RESIZE_PULL_SCROLLBACK`, input type `bool*`.
- The GhosttyVtScreen ctor sets it to false under `#if PLATFORM_WINDOWS` and ignores the result.

**6.** ProbeLoad in GhosttyNative.cs now accepts `QUICKSHEET_LIBGHOSTTY` only if `Path.IsPathFullyQualified`, on every OS. It skips the bare-name probe on Windows and keeps it on Linux and macOS.

**8.** Terminal/LegacyVtScreen.cs:
- `_current` is capped at the last 64 KiB.
- A new `IncompleteEscape` regex (`\x1b(?:\[[0-9;?]*|\][^\x07]*|\(?)$`) is applied to the partial line only, before the main strip.
- Scratch test results: after "ab\x1b[3" GetText is `ab`. After "1mred\x1b[0m\n" it is `abred`. A 200 KiB line with no newline gives GetText length 65536.

**9/10.**
- GetText decodes with `Marshal.PtrToStringUTF8(ptr, checked((int)len))`.
- A `_formatFailures` counter sets `_dirty = false` after 3 consecutive format failures, still returning the cache. The counter resets on success.

**11.**
- The two identical catch blocks in `GhosttyNative.Init` are merged into one `catch (Exception ex)`.
- `VtScreen.UsingGhostty` is removed (grep found no uses).
- CLAUDE.md now says "Scrollback capped at ~5,000 lines (page-granular in libghostty)."
- README.md line 50 now reads `powershell -ExecutionPolicy Bypass -File scripts/build-libghostty.ps1`. The .ps1 header (line 9) already said the same, so I left it alone.

**Windows compile check:** a scratch net9.0 project under the scratchpad (`.../scratchpad/chk/win`, `PLATFORM_WINDOWS`, ConPtyProcess.cs, InlineProcessManager.cs, Terminal/*.cs, plus a `CellPrefix.IsCommand` stub) builds with 0 warnings and 0 errors.

**Deviations:** none. `using System.Text;` is now unused in GhosttyVtScreen.cs but harmless, so I left it.
