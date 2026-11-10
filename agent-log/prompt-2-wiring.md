# Agent 2 — wire PTY readers to VtScreen (Sonnet)

Exact prompt given by the orchestrator (`$SCRATCH` = session scratch dir, outside the repo).

````text
You are editing C# in the repo at /home/user/QuickSheet (a .NET 9 desktop-wallpaper spreadsheet; read /home/user/QuickSheet/CLAUDE.md first). Do NOT commit. Only edit these three files: /home/user/QuickSheet/ConPtyProcess.cs, /home/user/QuickSheet/Platform/Linux/PtyProcess.cs, /home/user/QuickSheet/InlineProcessManager.cs. Other agents are concurrently creating files under /home/user/QuickSheet/Terminal/ and editing docs/scripts/csproj — don't touch those. Zero NuGet packages; keep hand-written P/Invoke style.

## Context
Inline cells (`i:` cells referencing `r:` commands) run a subprocess in a pseudo-terminal (ConPTY on Windows in ConPtyProcess.cs, openpty+posix_spawn on Linux in PtyProcess.cs) and show its output. Today each reader thread decodes UTF-8 per 4 KB chunk (corrupts multibyte chars split across reads), splits on '\n', strips ANSI with a duplicated regex, and (Linux only) does buggy '\r' handling. Cursor movement, erase-line, clear-screen, alt-screen TUIs, progress bars all render wrong.

We're migrating this console emulation to libghostty-vt (Ghostty's VT emulator). Another agent is writing this API in namespace `ExcelConsole.Terminal` (files Terminal/IVtScreen.cs, Terminal/VtScreen.cs, etc.). Code against EXACTLY this contract (it may not exist on disk yet when you start; if it doesn't, create NOTHING in Terminal/ — just code against it; you can check later whether it appeared to build):

```csharp
namespace ExcelConsole.Terminal;
internal interface IVtScreen : IDisposable
{
    void Write(ReadOnlySpan<byte> data);   // raw PTY bytes; thread-safe
    void Resize(int cols, int rows);       // thread-safe
    bool IsDirty { get; }                  // written since last GetText()
    string GetText();                      // scrollback+screen plain text, '\n'-joined, trailing blank lines removed, "" if empty; clears IsDirty; cached
}
internal static class VtScreen
{
    public static IVtScreen Create(int cols, int rows, int maxScrollbackLines); // libghostty if available, else legacy regex fallback
    public static bool UsingGhostty { get; }
}
```

## Required changes

### ConPtyProcess.cs (Windows, inside `#if PLATFORM_WINDOWS`)
- Keep all ConPTY P/Invoke/startup/cleanup logic as is.
- Replace `_outputLines`, `_lock`, line buffering, `AppendLine`, `AnsiPattern`/`StripAnsi`, `System.Text.RegularExpressions` using with a single `private IVtScreen? _screen;` created in Start() with `VtScreen.Create(cols, rows, MaxLines)` BEFORE the reader thread starts (create it at the top of Start so failure paths can dispose it, or create it right before starting the thread — your call, but make sure it's disposed on every failure path and in Dispose AFTER the reader thread is joined).
- ReadOutput: read bytes and call `_screen.Write(buffer.AsSpan(0, n))`. No string decoding.
- `HasNewOutput` → `_screen?.IsDirty ?? false` (read-only property now; check nothing else sets it — grep the repo).
- `GetOutput()` → `string? GetOutput() { var t = _screen?.GetText(); return string.IsNullOrEmpty(t) ? null : t; }`
- Add `public void Resize(int cols, int rows)` that calls `ResizePseudoConsole(_hPC, new COORD{...})` (add the kernel32 P/Invoke: `int ResizePseudoConsole(IntPtr hPC, COORD size)`) and `_screen?.Resize(cols, rows)`. Clamp to [1, short.MaxValue].
- Thread-safety on dispose: the reader thread may be in Write while Dispose runs; IVtScreen is thread-safe and Write after Dispose is a no-op, but still dispose the screen only after `_readerThread?.Join(2000)`.

### Platform/Linux/PtyProcess.cs (inside `#if PLATFORM_LINUX`)
- Same treatment: `IVtScreen? _screen` created in Start() via `VtScreen.Create(cols, rows, MaxLines)`; ReadLoop writes raw bytes `_screen.Write(buf.AsSpan(0, (int)n))`. Remove the regex, the StringBuilder line logic, the Regex using, `_outputLines`, `ReplaceLastLine`.
- The exit marker: today ReadLoop appends a line `[exited N]` after the child ends. Keep that behaviour but do NOT feed it into the emulator (a TUI may have left the cursor mid-screen). Store it in a `volatile string? _exitMarker` and have GetOutput return `text` + ("\n" if text non-empty) + marker. Make sure setting it marks output as new: keep a `volatile bool _markerPending` that HasNewOutput ORs in, cleared by GetOutput.
- `CreateError(string message)`: no child process; must produce output == message and HasExited == true. Implement with a field `_errorText` (or reuse the marker mechanism) instead of a screen — simplest correct approach, no native allocation.
- `HasNewOutput` → `(_screen?.IsDirty ?? false) || _markerPending`.
- `Resize(cols, rows)`: keep the ioctl, and also call `_screen?.Resize(cols, rows)`.
- Dispose: after closing the master fd and joining the reader thread, dispose the screen.
- ALSO: TERM is set to `xterm-256color` in the child env. Keep it (libghostty is xterm-compatible).

### InlineProcessManager.cs
- Update the class doc comment: PTY output is now interpreted by a real VT emulator (libghostty-vt via `ExcelConsole.Terminal.VtScreen`, regex fallback when the native lib is absent), so cursor movement / clear / progress bars / TUIs render as the final screen + scrollback.
- `MaxOutputLines` const is currently unused; leave it or pass it through — your call, but don't leave two different constants that disagree (ConPtyProcess.MaxLines / PtyProcess.MaxLines are 5000 too; fine to keep them).
- Optional, only if trivial and safe: when `EnsureRunning` is called with the same command but different ptyCols/ptyRows than last time, call a new `ManagedProcess.Resize(cols, rows)` that forwards to the pty's Resize (track the last size in ManagedProcess). This lets the emulator reflow when the span size changes. Keep the 4096-col JSON case working (cols up to 4096 must be fine; ConPTY COORD is short so 4096 is OK).
- Add `using ExcelConsole.Terminal;` only where needed.

## Verification
.NET 9 SDK: $SCRATCH/tools/dotnet/dotnet (not on PATH). This container is Linux, so `$DOTNET build /home/user/QuickSheet/ExcelConsole.csproj` only compiles the Linux path. Once Terminal/*.cs exists (another agent is writing it; poll by checking for /home/user/QuickSheet/Terminal/VtScreen.cs a few times, it should appear within ~10 minutes — if it never appears, write a temporary stub ONLY inside a scratch project, never in the repo), build and fix any errors in YOUR files. To compile-check the Windows ConPtyProcess.cs on Linux, create a throwaway project in $SCRATCH/wincheck (NOT in the repo): OutputType Library, TargetFramework net9.0, `<DefineConstants>$(DefineConstants);PLATFORM_WINDOWS</DefineConstants>`, Nullable/ImplicitUsings enable, EnableDefaultCompileItems false, and `<Compile Include="/home/user/QuickSheet/ConPtyProcess.cs;/home/user/QuickSheet/InlineProcessManager.cs;/home/user/QuickSheet/Terminal/*.cs;/home/user/QuickSheet/CellPrefix.cs" />` (add more repo files only if needed for CellPrefix dependencies; if CellPrefix drags in too much, add a minimal stub of `CellPrefix.IsCommand` in the scratch project instead). Make it build. Also, if possible on Linux, write a quick scratch console test that uses PtyProcess (compile PtyProcess.cs + Terminal/*.cs with PLATFORM_LINUX) to run `printf 'héllo\r\nprogress 10%%\rprogress 99%%\r\n\033[31mred\033[0m\n'` and print GetOutput() after exit — expect "héllo", "progress 99%", "red", "[exited 0]" (legacy fallback backend will be used since the native lib is missing here). Also test CreateError("boom") → GetOutput()=="boom", HasExited==true.

Report back: summary of the diff per file, test outputs, and anything you were unsure about.
````
