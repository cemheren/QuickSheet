All three files are edited and building. Nothing is committed. The Linux repo build (`ExcelConsole.csproj`) has 0 errors and 0 warnings. A scratch Windows-define build of ConPtyProcess.cs, InlineProcessManager.cs and Terminal/*.cs also gives 0 errors and 0 warnings. Both builds ran against the real Terminal/*.cs (a stub was used before it appeared).

**ConPtyProcess.cs**
- Removed the regex and `Text` usings, `_outputLines`, `_lock`, `AppendLine`, `AnsiPattern` and `StripAnsi`.
- Added `private IVtScreen? _screen`.
- Added the `ResizePseudoConsole(IntPtr, COORD)` P/Invoke and `public void Resize(cols, rows)`, clamped to [1, short.MaxValue]; calls the P/Invoke only if `_hPC` is non-zero, then `_screen?.Resize`.
- `Start()` clamps cols/rows and creates the screen right before the reader thread starts, after `CreateProcessW` has succeeded (no failure path after that point, so no leak).
- `ReadOutput` now only calls `_screen?.Write(buffer.AsSpan(0, n))`.
- `HasNewOutput` is read-only: `_screen?.IsDirty ?? false`. A repo grep found nothing else that sets it.
- `GetOutput()` returns null when the text is empty.
- `Dispose` disposes the screen after `_readerThread?.Join(2000)`, then nulls it.

**Platform/Linux/PtyProcess.cs**
- Removed the regex and `Text` usings, `_outputLines`, `_lock`, `_hasNewOutput`, `AppendLine`, `ReplaceLastLine`, `AnsiPattern`, `StripAnsi`.
- Added `IVtScreen? _screen`, created after a successful spawn and before the reader thread; `ReadLoop` writes raw bytes to it.
- Added `volatile string? _exitMarker`, `volatile bool _markerPending`, `string? _errorText`.
- On child exit, `ReadLoop` sets `_exitMarker = "[exited N]"` and `_markerPending = true`. The marker is never fed to the emulator.
- `GetOutput()` clears `_markerPending` first (so a concurrent marker re-arms it), returns `_errorText` if set, else screen text + "\n" (if non-empty) + marker.
- `CreateError` sets `_exited`, `_errorText`, `_markerPending`; no native resources.
- `HasNewOutput` = `(_screen?.IsDirty ?? false) || _markerPending`.
- `Resize` calls `_screen?.Resize` then the ioctl.
- `Dispose` disposes the screen after closing the master fd and joining the reader.

**InlineProcessManager.cs**
- Updated class doc comment.
- Added `LastCols`/`LastRows` and `ManagedProcess.Resize(cols, rows)`, forwarding only on change; set at start on both platforms.
- `EnsureRunning` with the same command now calls `existing.Resize(ptyCols, ptyRows)` before returning. 4096 cols is fine.
- `MaxOutputLines` still unused (5000, agrees with both `MaxLines`).

**Test** (scratch console, PtyProcess.cs + Terminal/*.cs, PLATFORM_LINUX, legacy backend since libghostty-vt isn't in the container) on `printf 'héllo\r\nprogress 10%%\rprogress 99%%\r\n\033[31mred\033[0m\n'`:
```
héllo
progress 99%
red
[exited 0]
```
`HasNewOutput` true before `GetOutput()`, false after. `CreateError("boom")` → `GetOutput()=="boom"`, `HasExited==true`, `HasNewOutput==false` after read. `Resize`/`Dispose` on the error placeholder don't throw.

**Unsure about**
- Windows path is only compile-checked, not run.
- PtyProcess `Dispose` joins the reader with a 500 ms timeout as before; if the reader is still in `read()`, the screen may be disposed while it calls `Write`. Relies on `Write` after `Dispose` being a no-op (it is, per contract).
