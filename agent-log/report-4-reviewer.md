# Reviewer report (Opus), on the first feature commit

## Review of HEAD 31c757d ("Migrate inline-process console emulation to libghostty-vt")

I made no edits to the repo. All builds and experiments ran on copies under the scratchpad.

### Findings, most severe first

**1. BLOCKER: a fresh Windows clone no longer builds.** `ExcelConsole.csproj:53`
- **What's wrong:** `<None Include="native/win-$(QsNativeArch)/ghostty-vt.dll" CopyToOutputDirectory=...>` is a literal path, not a wildcard. MSBuild keeps a literal item even when the file is missing, and the copy step then fails.
- **Verified:** I reproduced it in a scratch csproj with the same item and got `error MSB3030: Could not copy the file ".../native/win-x64/ghostty-vt.dll" because it was not found.`
- **Failure scenario:** anyone on Windows who hasn't run the Zig script gets a failed `dotnet build`. That breaks "clone, build, run" and the "app must run without them" rule the commit adds to CLAUDE.md and CONTRIBUTING.md.
- **Linux is fine** because its line (`:58`) uses the wildcard `libghostty-vt.so*`. A HEAD copy built cleanly on Linux.
- **Fix:** add `Condition="Exists('native/win-$(QsNativeArch)/ghostty-vt.dll')"`, or use a wildcard such as `native/win-$(QsNativeArch)/ghostty-vt*.dll`.

**2. MAJOR: the PTY and emulator are resized back and forth on every frame.** `InlineProcessManager.cs:100-104`, `:33-39`
- **What's wrong:** `EnsureRunning` now resizes whenever the requested size differs from the last one. But one process key is requested with different sizes within a single frame.
- **Linux trace** (`Platform/Linux/DesktopWindow.cs`):
  - The span pre-scan (`:612`) calls `ResolveDynamicValue(r,c, ptyCols, ptyRows)` with the span size, e.g. 39x4.
  - The span anchor is deliberately not occluded ("Anchor stays renderable"), so the row loop (`:679`) resolves the same cell again via `ResolveDynamicValue(r,c)`, which defaults to 120x30.
  - The status bar with F1 (`:904`) also uses 120x30.
- **Windows is identical:** pre-scan `DesktopForm.cs:505` vs row loop `:569`, status bar `:774`, Ctrl+Shift+C `:1117`.
- **Chains cause it too:** `commandOutput` is keyed by the cell that owns the process, but the size comes from the outermost cell (`DesktopWindow.cs:1872-1881`, `DesktopForm.cs:338-348`).
  - A `j:` cell pointing at an `i:` cell runs that `i:` process at 4096 cols, while the `i:` cell's own render asks for 120 or its span size. It flips every frame.
  - `i:` → `i:` chains with different spans do the same.
- **Failure scenario:** every 100 ms render (Linux) or 500 ms timer (Windows), each spanning `i:` cell gets:
  - two `TIOCSWINSZ` calls (two SIGWINCHs) or two `ResizePseudoConsole` calls;
  - two ghostty reflows of up to 5000 lines of scrollback;
  - two full reformats.
- **Knock-on effects:**
  - `Resize` marks the screen dirty, so `HasAnyNewOutput` never goes quiet, even after the child has exited. The wallpaper burns CPU at 10 fps indefinitely.
  - A TUI (htop, watch) redraws for one size while the emulator is already at the other, so the span shows garbled output.
  - The old code ignored size for an existing process, so none of this happened before.
- **Fix:** only the span pre-scan should drive size.
  - Make `ptyCols`/`ptyRows` nullable on `EnsureRunning` and the `ResolveDynamicValue` wrappers, and skip `Resize` when they are null. Pass null from the row loop, status bar and copy paths, or have the row loop reuse the anchor content already collected in `inlineSpans`.
  - For chains, resize only when `(processRow, processCol)` is the cell that started the resolve.

**3. MAJOR: a single-cell `i:` can now show an empty value.**
- **Where:** `Terminal/GhosttyVtScreen.cs:306`, consumed at `DesktopWindow.cs:679-681` (`displayVal[..nl]`) and the equivalent path in `DesktopForm.cs`.
- **What's wrong:** the old code dropped empty lines, so a single cell showed the first non-empty line. The ghostty plain formatter emits leading blank rows as `\n` (checked in `formatter.zig` around lines 1303-1332), and `GetText` only trims the end.
- **Failure scenario:** on Windows, `r: ping 8.8.8.8 -t` and `r: ipconfig` both start their output with `\r\n`, so the single-cell `i:` now displays blank. Before, it showed "Pinging …" or "Windows IP Configuration". The same happens for any Linux command that prints a leading newline.
- **Fix:** add `TrimStart('\n')` in `GetText`, or have single-cell callers pick the first non-empty line.

**4. MAJOR (exact threshold uncertain): large JSON output for `j:` cells gets truncated by the scrollback byte cap.** `GhosttyVtScreen.cs:220-221`, together with the 4096-col JSON size (`DesktopWindow.cs:1875`, `DesktopForm.cs:341`)
- **What's wrong:** a ghostty `Cell` is a `packed struct(u64)` (8 bytes, `page.zig:2389`), and pages hold full-width rows. At 4096 cols that is about 32 KB per row, so `SCROLLBACK_MAX_BYTES = 64 MB` keeps only about 2,000 physical rows, well under the intended 5,000.
- **Failure scenario:** pretty-printed JSON longer than about 2k lines (e.g. `kubectl get pods -o json`, `gh api`) loses its head, and `SelectJson` fails to parse it. Each such cell can also hold up to 64 MB.
- **Fix:** derive the byte cap from the size, e.g. `cols * 8 * maxLines` plus slack. Or, since UNWRAP rejoins soft-wrapped lines anyway, use a narrower terminal for JSON (e.g. 512 cols), noting that wrapped rows then count against the physical line limit.

**5. MINOR: the ConPTY resize behaviour recommended by the header is not set.** `GhosttyVtScreen.cs` constructor
- `GHOSTTY_TERMINAL_OPT_RESIZE_PULL_SCROLLBACK = 41` (`bool*`, in the 1.0.0 `terminal.h:2530-2555`) says to set it to false for Windows ConPTY. Otherwise the emulator and conhost disagree about screen contents after a resize, and this commit now resizes ConPTY (very often, given finding 2).
- **Fix:** under `PLATFORM_WINDOWS`, call `ghostty_terminal_set(t, 41, &false)` (a 1-byte bool).

**6. MINOR: Windows DLL-search hijack risk.** `GhosttyNative.cs:182-186`
- The last fallback, `NativeLibrary.TryLoad("ghostty-vt.dll")`, uses the default `LoadLibrary` search, which includes PATH and the current directory. A `ghostty-vt.dll` planted in a writable PATH directory or the working directory would be loaded into the wallpaper process.
- **Fix:** drop the bare-name probe on Windows, or use `TryLoad(name, assembly, DllImportSearchPath.AssemblyDirectory | DllImportSearchPath.System32, out h)`.

**7. MINOR: the native library is picked by the build machine's architecture, not the target's.** `ExcelConsole.csproj:47-49`
- `QsNativeArch` comes from the host's `OSArchitecture`. Publishing with `-r linux-arm64` from an x64 machine silently copies the x64 `.so`, which then fails to load and falls back to legacy. X86 hosts produce a `native/win-X86/` path that never exists.
- **Fix:** derive the architecture from `$(RuntimeIdentifier)` when it is set.

**8. MINOR: `LegacyVtScreen` cost grows without bound.** `LegacyVtScreen.cs:465`
- `GetText` now runs the regex `Strip` over the partial line `_current` on every dirty call, and `_current` has no size limit. A stream with no `\n` (e.g. ConPTY repaints using cursor positioning) makes each frame cost O(total bytes).
- **Fix:** cap `_current` (e.g. keep the last 64 KB) and cache the stripped partial line.

**9. NIT: `GhosttyVtScreen.cs:293-295` copies the formatted text one extra time.** It goes native buffer → `byte[]` → `string`. Decode straight from the pointer instead (`Marshal.PtrToStringUTF8(ptr, (int)len)` or `Encoding.UTF8.GetString((byte*)ptr, (int)len)`). The result is often larger than 85 KB, so it lands on the large object heap on every dirty frame.

**10. NIT: a persistent format failure re-renders forever.** If `ghostty_formatter_format_alloc` keeps failing (`GhosttyVtScreen.cs:300-304`), the screen stays dirty and the render loop never idles. Consider clearing `_dirty` after N failures.

**11. NIT: small cleanups.**
- `GhosttyNative.cs:144-154` has two catch blocks that do the same thing.
- `VtScreen.UsingGhostty` is unused.
- `LegacyVtScreen` briefly shows `[3` for an escape sequence split across reads (I confirmed this with a scratch run).
- The CLAUDE.md line "Output capped at 5,000 lines" now means physical rows plus a byte cap.
- README suggests running `scripts/build-libghostty.ps1` directly; the default ExecutionPolicy blocks that. Use the `powershell -ExecutionPolicy Bypass -File …` form from the script header.

**12. Untested: the libghostty code path has never run.** The commit's own plan log says the native library was never built, so `GhosttyVtScreen` and `GhosttyNative` are checked only against the headers. Before merging, run a smoke test against a real build: start, format, resize, and dispose while the reader thread is writing.

### Checked and found correct
- **P/Invoke signatures** match the 1.0.0 headers exactly:
  - `terminal_new(const GhosttyAllocator*, GhosttyTerminal*, uint16_t, uint16_t)`
  - `terminal_free`
  - `vt_write(GhosttyTerminal, const uint8_t*, size_t)` (`ref byte` pins the buffer; `data` is never empty there)
  - `resize(GhosttyTerminal, u16, u16, u32, u32)`
  - `terminal_set` / `formatter_set` with `const void*` pointing to the value
  - `formatter_terminal_new(allocator, out, terminal)`
  - `format_alloc(formatter, allocator, uint8_t**, size_t*)`
  - `ghostty_free(allocator, ptr, len)` with a NULL allocator on both allocate and free
  - `ghostty_version()` returns a static `const char*` that is never freed
  - Results are `int` (enums are backed by `c_int`); the convention is `.c`, i.e. Cdecl, which is a no-op on x64.
- **Option values:**
  - `SCROLLBACK_MAX_BYTES = 28` and `MAX_LINES = 29`, both `size_t*` (written as pointer width).
  - Formatter `EMIT = 1` with a 4-byte enum, `PLAIN = 0`; `UNWRAP = 2` and `TRIM = 3` with 1-byte bools.
  - Empty output (Success with a NULL pointer) is handled.
  - The default formatter selection covers scrollback plus the active screen.
- **Lifetime and threading:**
  - The formatter is freed before the terminal, and the constructor's failure path frees both.
  - All native calls run under `_lock` with a `_disposed` check, so a `Write` after `Dispose` (reader still running after the `Join` timeout) is safe and there is no use-after-free.
  - The `_markerPending` / `_exitMarker` ordering (clear first, then read) re-arms correctly.
  - The `GetText` cache and dirty-flag logic is correct.
- **Library loading:**
  - One `SetDllImportResolver` per assembly, registered lazily. It returns `IntPtr.Zero` for any other library, so X11, libutil and kernel32 imports fall back to default probing.
  - Probe order is the env var, then the exe directory, then the system. The major-version check matches what `version.h` promises.
- **LegacyVtScreen** (scratch run): `\r\n` split across reads, `\r` progress overwrite, multi-byte UTF-8 split across reads, and the line cap all behave correctly. `GetCharCount` does not change decoder state.
- **Build scripts:**
  - Both pinned SHA-256 values match the official `release.files.ghostty.org/libghostty-vt/1.0.0/SHA256SUMS`.
  - A SHA mismatch stops the build in both scripts, and a minisign failure stops it too.
  - Temp directories are cleaned up via `trap` and `finally`, and the Zig cache starts fresh and empty, as the packaging doc requires.
  - Flags `--system`, `-Demit-lib-vt` (redundant for a tarball but harmless), `-Dcpu=baseline`, `-Dstrip=true` and Zig 0.17.x match PACKAGING-LIBGHOSTTY.md.
  - The Windows DLL is taken from `bin/ghostty-vt.dll`; the Linux file is the SONAME `.so.1`, dereferenced with `cp -L`.
- **csproj:** the Linux `libghostty-vt.so*` link keeps the `.so.1` name. The `Compile`/`None Remove` of `native/**` and `scripts/**` runs before the native `Include`, so it doesn't drop the library item.
- **Builds:** HEAD builds on Linux, and `ConPtyProcess.cs` plus `Terminal/*.cs` compile with `PLATFORM_WINDOWS` defined. `ResizePseudoConsole(IntPtr, COORD)` is declared correctly.

Scratch work is under `$SCRATCH/` in the `review-msb`, `review-copy`, `review-win`, `review-legacy` and `review-sums` folders.
