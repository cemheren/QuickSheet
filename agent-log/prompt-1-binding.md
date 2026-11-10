# Agent 1 — libghostty-vt P/Invoke + VtScreen (Sonnet)

Exact prompt given by the orchestrator (`$SCRATCH` = session scratch dir, outside the repo).

````text
You are writing C# code in the repo at /home/user/QuickSheet (a .NET 9 desktop-wallpaper spreadsheet; read /home/user/QuickSheet/CLAUDE.md first). Do NOT commit, do NOT touch any files other than the ones listed below — other agents are editing other files in parallel. HARD RULE: zero NuGet packages; hand-written P/Invoke only (use [DllImport], matching repo style; do not enable unsafe code / AllowUnsafeBlocks).

## Goal
We're replacing the hand-rolled ANSI-regex "console emulation" for inline-process cells with libghostty-vt (Ghostty's VT emulator C library, stable C ABI v1.0, tag vt-1.0.0). You write the managed binding + a small screen abstraction. Another agent will wire PTY readers to it using exactly the contract below, so match it EXACTLY.

## Files to create (namespace `ExcelConsole.Terminal`, folder `/home/user/QuickSheet/Terminal/`)

### 1. Terminal/IVtScreen.cs
```csharp
namespace ExcelConsole.Terminal;
/// Terminal screen model fed with raw PTY output bytes. All members thread-safe
/// (PTY reader thread writes; UI thread reads).
internal interface IVtScreen : IDisposable
{
    /// Feed raw bytes exactly as read from the PTY (may split UTF-8 / escape sequences across calls).
    void Write(ReadOnlySpan<byte> data);
    /// Resize the emulated grid (cells). Values < 1 are clamped to 1, > 65535 clamped to 65535.
    void Resize(int cols, int rows);
    /// True if Write/Resize happened since the last GetText() call.
    bool IsDirty { get; }
    /// Plain text of scrollback + visible screen, lines joined with '\n', no '\r',
    /// trailing whitespace per line trimmed, trailing empty lines removed. "" when empty.
    /// Clears IsDirty. Must be cheap when not dirty (return cached string).
    string GetText();
}
```

### 2. Terminal/VtScreen.cs — factory
```csharp
internal static class VtScreen
{
    /// Creates a libghostty-backed screen when the native library is loadable, otherwise LegacyVtScreen.
    public static IVtScreen Create(int cols, int rows, int maxScrollbackLines);
    /// True if libghostty-vt was loaded (evaluated once, lazily).
    public static bool UsingGhostty { get; }
}
```
If GhosttyVtScreen construction throws (e.g. ghostty_terminal_new fails), fall back to LegacyVtScreen. Write a one-time diagnostic line to stderr (Console.Error.WriteLine) saying which backend is used and the libghostty version string (or why it couldn't load) — only once per process.

### 3. Terminal/GhosttyNative.cs — raw P/Invoke (internal static class GhosttyNative)
Library name const `Lib = "ghostty-vt"`. All imports `CallingConvention = CallingConvention.Cdecl`, `ExactSpelling = true`. C ABI facts (verified from headers at tag vt-1.0.0):
- Handles `GhosttyTerminal`, `GhosttyFormatter` are opaque pointers → IntPtr.
- `GhosttyResult` is a C enum (int): SUCCESS=0, OUT_OF_MEMORY=-1, INVALID_VALUE=-2, OUT_OF_SPACE=-3, UNSUPPORTED=-8.
- `GhosttyResult ghostty_terminal_new(const GhosttyAllocator* allocator /*pass IntPtr.Zero = default*/, GhosttyTerminal* terminal, uint16_t cols, uint16_t rows);` (cols/rows must be > 0)
- `void ghostty_terminal_free(GhosttyTerminal t);`
- `void ghostty_terminal_vt_write(GhosttyTerminal t, const uint8_t* data, size_t len);` (never fails). Bind data as `ref byte` and pass `ref MemoryMarshal.GetReference(span)`; skip the call when span is empty. size_t → nuint.
- `GhosttyResult ghostty_terminal_resize(GhosttyTerminal t, uint16_t cols, uint16_t rows, uint32_t cell_width_px, uint32_t cell_height_px);` (pass e.g. 8 and 16 for px; only used for image protocols)
- `GhosttyResult ghostty_terminal_set(GhosttyTerminal t, int option, const void* value);` value is a POINTER to the value for non-pointer types.
   - GHOSTTY_TERMINAL_OPT_SCROLLBACK_MAX_BYTES = 28 (input: size_t*; default is only ~10,000 bytes ≈ one ~400KB page — too small, set it to e.g. 64 MiB)
   - GHOSTTY_TERMINAL_OPT_SCROLLBACK_MAX_LINES = 29 (input: size_t*; page-granular estimate)
   Pass values via a small helper that allocates with Marshal.AllocHGlobal(IntPtr.Size), writes nuint, calls, frees (or stackalloc-free approach w/o unsafe: Marshal.AllocHGlobal is fine).
- `GhosttyResult ghostty_formatter_terminal_new(const GhosttyAllocator* allocator, GhosttyFormatter* formatter, GhosttyTerminal terminal);` Formatter borrows terminal; terminal must outlive formatter. With no selection set it formats the WHOLE page list of the active screen = scrollback + visible area (primary screen); on alt screen just the visible screen. That's what we want.
- `GhosttyResult ghostty_formatter_set(GhosttyFormatter f, int option, const void* value);`
   - GHOSTTY_FORMATTER_OPT_EMIT = 1 (value: GhosttyFormatterFormat, a C int enum; PLAIN = 0, VT = 1, HTML = 2) — set PLAIN explicitly.
   - GHOSTTY_FORMATTER_OPT_UNWRAP = 2 (value: C `bool` = 1 byte! pass pointer to a single byte 1/0) — set TRUE so soft-wrapped lines are rejoined (consumers re-wrap themselves and JSON parsing needs unbroken lines).
   - GHOSTTY_FORMATTER_OPT_TRIM = 3 (C bool, 1 byte) — set TRUE.
- `GhosttyResult ghostty_formatter_format_alloc(GhosttyFormatter f, const GhosttyAllocator* allocator, uint8_t** out_ptr, size_t* out_len);` → `out IntPtr ptr, out nuint len`. Output is UTF-8; free with ghostty_free. ptr may be null/len 0 when empty.
- `void ghostty_formatter_free(GhosttyFormatter f);`
- `void ghostty_free(const GhosttyAllocator* allocator, uint8_t* ptr, size_t len);` (pass IntPtr.Zero allocator — must match the allocator used for format_alloc, which is also NULL/default).
- `const char* ghostty_version(void);` → returns IntPtr; read with Marshal.PtrToStringUTF8. Returns e.g. "1.0.0".

Library loading — implement `public static bool IsAvailable` (lazy, computed once, thread-safe via Lazy<T>) and `public static string? Version` / `public static string? LoadError`:
- Register `NativeLibrary.SetDllImportResolver(typeof(GhosttyNative).Assembly, Resolver)` exactly once inside the Lazy init (guard it; SetDllImportResolver throws if called twice for the same assembly — catch InvalidOperationException). IMPORTANT: no other code in this assembly registers a resolver today; keep it that way but be defensive.
- Resolver only handles libraryName == Lib; returns IntPtr.Zero for others.
- Probe order: (1) env var `QUICKSHEET_LIBGHOSTTY` = full path to the library file; (2) files in `AppContext.BaseDirectory`: Windows `ghostty-vt.dll`; Linux `libghostty-vt.so.1`, then `libghostty-vt.so`; macOS `libghostty-vt.1.dylib`, `libghostty-vt.dylib`; (3) bare names via NativeLibrary.TryLoad(name) so a system-installed lib (e.g. /usr/lib/libghostty-vt.so.1 via the dynamic loader) is found. Cache the loaded handle.
- After load, call ghostty_version() and require the major version == 1 (parse up to first '.'); otherwise IsAvailable=false with LoadError explaining. Wrap everything in try/catch (DllNotFoundException, EntryPointNotFoundException, BadImageFormatException) → IsAvailable=false.
- Add a doc comment at top explaining: libghostty-vt is NOT a NuGet package; it's a native lib built from a pinned, checksum-verified source tarball by scripts/build-libghostty.sh / .ps1 and copied next to the executable (native/ folder → output dir). If absent, QuickSheet falls back to LegacyVtScreen.

### 4. Terminal/GhosttyVtScreen.cs — `internal sealed class GhosttyVtScreen : IVtScreen`
- ctor(int cols, int rows, int maxScrollbackLines): clamp cols/rows to [1, 65535]; ghostty_terminal_new; set SCROLLBACK_MAX_LINES = maxScrollbackLines and SCROLLBACK_MAX_BYTES = 64 MiB; create formatter and set EMIT=PLAIN, UNWRAP=true, TRIM=true. Throw InvalidOperationException with the result code on any non-success (and free what was created).
- One private `_lock` object guards every native call (libghostty terminal is not thread-safe; reader thread writes, UI thread formats).
- Write: under lock call vt_write, set dirty.
- Resize: under lock call ghostty_terminal_resize(cols, rows, 8, 16); set dirty. Skip if same size as current.
- GetText: under lock; if !dirty and cache != null return cache; else format_alloc → copy to managed byte[] (Marshal.Copy) → ghostty_free → Encoding.UTF8.GetString → normalize: remove '\r', TrimEnd('\n', ' ') per the contract (trailing empty lines removed). Cache, clear dirty, return.
- Dispose: idempotent; under lock free formatter FIRST, then terminal; set handles to zero. After dispose, Write is a no-op and GetText returns the last cached text (or "").
- Use a finalizer? No — keep it simple; IDisposable is enough (owners always dispose). Do mark the class sealed.

### 5. Terminal/LegacyVtScreen.cs — `internal sealed class LegacyVtScreen : IVtScreen` (fallback when the native lib is missing)
Move the existing behaviour here, de-duplicated (currently copy-pasted in /home/user/QuickSheet/ConPtyProcess.cs and /home/user/QuickSheet/Platform/Linux/PtyProcess.cs — read both). Keep the same regex:
`@"\x1b(?:\[[0-9;?]*[A-Za-z@~]|\][^\x07]*\x07|\([A-Z0-9]|[A-Z=<>7-9])|[\x00-\x08\x0e-\x1f]"` (RegexOptions.Compiled).
Behaviour: decode bytes with a persistent `System.Text.Decoder` from Encoding.UTF8.GetDecoder() (so multi-byte chars split across Write calls are not corrupted — that's a real bug in the old code). Then the Linux semantics: '\n' → strip ANSI from current line (TrimEnd '\r'), append if non-empty; bare '\r' → strip, ReplaceLastLine if non-empty... careful: the old Linux code replaced the LAST committed line on '\r', which is wrong for "\r\n" sequences arriving as "\r" then "\n"? Look closely: for "text\r\n" the old code hits '\r' first → ReplaceLastLine("text") (overwrites previous line!) then '\n' with empty buffer → nothing. That is a bug. Implement saner semantics instead: keep a "current line" buffer; on '\r' set a flag `_pendingCr=true`; on '\n' commit the current line (strip ANSI; append if non-empty) and clear flag; on any other char, if `_pendingCr` then clear the current line buffer first (carriage return overwrite = progress bars) and clear the flag. GetText returns committed lines + (stripped) current partial line if non-empty, joined with '\n'. Cap committed lines at maxScrollbackLines (remove from front; use a Queue/LinkedList or List with RemoveRange to avoid O(n^2)). Resize: no-op apart from dirty flag. Same IsDirty/caching/locking semantics as the contract.

## Verification
A .NET 9 SDK is installed at $SCRATCH/tools/dotnet/dotnet (not on PATH). Run `$DOTNET build /home/user/QuickSheet/ExcelConsole.csproj` and make it compile with zero new warnings from your files (other agents may be mid-edit on ConPtyProcess.cs / PtyProcess.cs / InlineProcessManager.cs — if errors are ONLY in those files, ignore them and say so). The native lib is NOT available in this container, so additionally write a tiny throwaway console project in $SCRATCH/legacytest (NOT in the repo) that compiles Terminal/*.cs via <Compile Include> links and exercises LegacyVtScreen: split UTF-8 across writes ("é" bytes split), "a\r\nb\r\n" → "a\nb", progress bar "10%\r20%\r30%\n" → "30%", ANSI colored text stripped, VtScreen.Create falls back to legacy and UsingGhostty==false. Run it and report results. Do not put that test project in the repo.

Report back: list of files created, the public surface, the test output, and any deviations from this spec (with reasons).
````
