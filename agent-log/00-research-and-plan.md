# libghostty-vt migration — research notes & plan (orchestrator)

Feature branch: `ccr-7ea14cd6-efxsg4` (PR against `master`).
This log branch: `ccr-7ea14cd6-efxsg4-agent-log` (documentation only, not for merge).

## What "console emulation" meant in QuickSheet before this change

`i:` cells that point at `r:` commands run the command in a pseudo-terminal and show its output:

| File | Role | Emulation it did |
|---|---|---|
| `ConPtyProcess.cs` (Windows) | ConPTY + `cmd.exe /c` | Decode each 4 KB read as UTF-8, split on `\n`, `TrimEnd('\r')`, strip ANSI with a regex, drop empty lines, cap 5000 lines |
| `Platform/Linux/PtyProcess.cs` | `openpty` + `posix_spawn("/bin/sh -c")`, `TERM=xterm-256color` | Same regex (copy-pasted), plus bare `\r` → "replace last committed line" (which also mis-handled `\r\n`) |
| `InlineProcessManager.cs` | Keyed by pointer cell, exposes `GetOutput()` string | — |

Consumers (`DesktopWindow.cs` / `DesktopForm.cs`) just take the string, re-wrap it to the span width and show the last N lines. JSON cells request 4096 columns and parse the text.

Bugs this caused: cursor movement (`CSI H`), erase (`CSI K/J`), clear-screen, alt-screen TUIs, and `\r` progress bars rendered as garbage or stale lines; multi-byte UTF-8 split across two reads got corrupted.

## libghostty findings

- libghostty-vt = Ghostty's VT parser + terminal state (screen, scrollback, styles, reflow) as a C library. Stable C ABI from **1.0** (release tag `vt-1.0.0`, commit `0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec`, 2026-11-06). SONAME `libghostty-vt.so.1`; Windows `bin/ghostty-vt.dll`.
- Functions we need, all present at `vt-1.0.0`:
  `ghostty_terminal_new(alloc, &t, u16 cols, u16 rows)`, `ghostty_terminal_free`, `ghostty_terminal_vt_write(t, u8*, size_t)` (never fails), `ghostty_terminal_resize(t, u16, u16, u32 px, u32 px)` (reflows primary screen), `ghostty_terminal_set(t, opt, const void*)`,
  `ghostty_formatter_terminal_new(alloc, &f, t)`, `ghostty_formatter_set`, `ghostty_formatter_format_alloc(f, alloc, u8**, size_t*)`, `ghostty_formatter_free`, `ghostty_free(alloc, ptr, len)`, `ghostty_version()`.
- Option values: `TERMINAL_OPT_SCROLLBACK_MAX_BYTES=28` (default only ~10 KB ⇒ one page; must raise), `TERMINAL_OPT_SCROLLBACK_MAX_LINES=29` (size_t*), `FORMATTER_OPT_EMIT=1` (PLAIN=0), `UNWRAP=2`, `TRIM=3` (C bool, 1 byte). Enums are `int`. Result: SUCCESS=0, negatives on error.
- Formatter with no selection formats the full page list of the **active** screen (`PageListFormatter(&screen.pages)` with no top-left bound) ⇒ scrollback + visible area on the primary screen, current frame only on the alt screen. Exactly the "final screen + history" semantics we want.
- Terminal is not thread-safe ⇒ one lock around all native calls per instance.
- Effects callbacks (write_pty for DSR replies etc.) intentionally **not** wired: Linux child stdin is `/dev/null` and the slave isn't its controlling tty, so replies can't reach it, and writing to the master with nobody reading could block the reader thread. Can be added later.
- Distribution: official signed source + deps tarballs at `release.files.ghostty.org/libghostty-vt/1.0.0/`, SHA256SUMS (minisign key `RWQlAjJC23149WL2sEpT/l0QKy7hMIFhYdQOFy0Z7z7PbneUgvlsnYcV`). Requires Zig 0.17.x. Offline build: `zig build --system <deps> -Doptimize=ReleaseFast -Dcpu=baseline`.
  - `libghostty-vt-1.0.0.tar.gz` sha256 `4a6ef758be721923e1969e7b81b443dcc92d8dd3fb5a39600c03a8817a753676`
  - `libghostty-vt-1.0.0-deps.tar.gz` sha256 `4d5ea5e448e4c7ea8fb55ccbaa8668f2a97d44f438cee831e2b7ecd162d85c29`
- No prebuilt Linux/Windows shared libraries are published (only xcframework + wasm).

## Policy fit

"Zero NuGet deps" is about supply chain. libghostty-vt is not a NuGet package; it is built from a pinned, checksum-verified official tarball by a repo script, never committed (`native/` gitignored), and bound with hand-written P/Invoke. To keep "clone, build, run" working without Zig, the app falls back to the old regex stripper (de-duplicated into `LegacyVtScreen`, with its UTF-8 and `\r\n` bugs fixed).

## Plan / delegation

Contract (written by orchestrator, handed to both coding agents so they could run in parallel on disjoint files):

```csharp
namespace ExcelConsole.Terminal;
internal interface IVtScreen : IDisposable {
    void Write(ReadOnlySpan<byte> data); void Resize(int cols, int rows);
    bool IsDirty { get; } string GetText();
}
internal static class VtScreen {
    public static IVtScreen Create(int cols, int rows, int maxScrollbackLines);
    public static bool UsingGhostty { get; }
}
```

| # | Agent (model) | Files |
|---|---|---|
| 1 | Binding (Sonnet) | `Terminal/IVtScreen.cs`, `VtScreen.cs`, `GhosttyNative.cs`, `GhosttyVtScreen.cs`, `LegacyVtScreen.cs` |
| 2 | Wiring (Sonnet) | `ConPtyProcess.cs`, `Platform/Linux/PtyProcess.cs`, `InlineProcessManager.cs` |
| 3 | Build + docs (Haiku) | `scripts/build-libghostty.{sh,ps1}`, `ExcelConsole.csproj`, `.gitignore`, README/CONTRIBUTING/CLAUDE.md/CHANGELOG/docs |
| 4 | Reviewer (Opus) | read-only review of the integrated commit |

## Verification limits in this environment

- .NET 9 SDK installed in the session scratchpad: repo builds with 0 warnings / 0 errors on Linux; Windows `ConPtyProcess.cs` compile-checked via a scratch project with `PLATFORM_WINDOWS`.
- Building libghostty-vt itself (running ghostty's Zig build) was **blocked by the session's permission policy** (executing downloaded third-party code), so the Ghostty backend path is verified only against the headers, not at runtime. The fallback path was exercised end-to-end through a real PTY.

## Orchestrator integration fixes (after agents, before review)

- `build-libghostty.sh`: copy from the stable SONAME symlink `lib/libghostty-vt.so.1` instead of hardcoding `libghostty-vt.so.1.0.0` (survives patch bumps).
- `CLAUDE.md`: corrected the pre-existing "pipe redirect on Linux" to "openpty + posix_spawn".

## Timeline

1. Research (orchestrator): repo survey, libghostty-vt headers at `vt-1.0.0`, packaging docs, official SHA256SUMS.
2. Agents 1–3 in parallel → integrated by orchestrator → commit `31c757d` "Migrate inline-process console emulation to libghostty-vt".
3. Agent 4 (reviewer) on `31c757d` → 1 blocker, 3 major, 4 minor, nits (see `report-4-reviewer.md`).
4. Agent 5 applied all fixes → orchestrator checked the platform-wrapper diff, removed an unused using → commit `534f184` "Address review findings on the libghostty-vt migration".
5. PR: https://github.com/cemheren/QuickSheet/pull/497 (not merged).

## Files in this folder

| File | What |
|---|---|
| `00-research-and-plan.md` | This file: research, decisions, delegation, timeline |
| `prompt-N-*.md` | Exact prompt each subagent received |
| `report-N-*.md` | Each subagent's final report |

Raw JSONL transcripts are deliberately not included: they carry session/account metadata and this repo is public. Ask if you want scrubbed copies.

## Phase 2

See `01-phase2-dependency-delivery.md`: the decision to take one dependency on libghostty-vt, NuGet research, the CI-built + pinned-download design, agents 6–11, the Windows-ABI catch, and the CI runs that published `libghostty-vt-1.0.0-1`.
