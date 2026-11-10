All five files are written, and the repo build (`ExcelConsole.csproj`) gives 0 warnings and 0 errors. The Ghostty path itself is untested, because the native library isn't in this container. Only the fallback was exercised.

**Files created** in `Terminal/`: `IVtScreen.cs`, `VtScreen.cs`, `GhosttyNative.cs`, `GhosttyVtScreen.cs`, `LegacyVtScreen.cs`. Nothing else in the repo was touched and nothing was committed.

**Public surface** (namespace `ExcelConsole.Terminal`):
- `IVtScreen` matches the contract exactly: `Write(ReadOnlySpan<byte>)`, `Resize(int, int)`, `IsDirty`, `GetText()`, plus `IDisposable`.
- `VtScreen.Create(cols, rows, maxScrollbackLines)` and `VtScreen.UsingGhostty`, as specified. The diagnostic line goes to stderr once per process and names the backend plus the libghostty version, or the reason it couldn't load.
- `GhosttyNative` has the `Lib` const, the raw imports with the exact signatures from the spec, and `IsAvailable`, `Version`, `LoadError` backed by a `Lazy<bool>`. It also has small helpers for passing values by pointer: `TerminalSetNuint`, `FormatterSetInt`, `FormatterSetBool`.
- `GhosttyVtScreen` and `LegacyVtScreen` are both `internal sealed` with ctor `(cols, rows, maxScrollbackLines)`.

**Test output** (throwaway project outside the repo):
```
[QuickSheet] terminal backend: legacy (libghostty-vt unavailable: native library not found (set QUICKSHEET_LIBGHOSTTY or place it next to the executable))
UsingGhostty=False type=LegacyVtScreen
PASS split utf8: [café]
PASS crlf: [a\nb]
PASS progress: [30%]
PASS ansi: [red ok]
PASS cap: [3\n4]
PASS dirty: [False]
```

**Deviations from the spec:**
- Explicit probe in `Init`: besides registering the resolver, `Init` also calls `ProbeLoad` directly and caches the handle, so detection still works if another resolver ever owns the assembly.
- Failed format call: if `format_alloc` fails, `GhosttyVtScreen.GetText` returns the previous cache and stays dirty so it retries.
- Skipped resize: if `ghostty_terminal_resize` fails, the stored size isn't updated.
- Legacy `Resize` only sets the dirty flag; the `cols`/`rows` ctor arguments are ignored.
- `LegacyVtScreen` clamps `maxScrollbackLines` to at least 1.
- Legacy path trims trailing spaces per line (not required by the spec).
- Legacy cap uses `List.RemoveRange`.
