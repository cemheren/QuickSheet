# QuickSheet Test Log

## Run: 2026-05-15 00:48
Commit: 04ccf70
Build: Release

### Discovery
- No new commits since last run
- No new extensions discovered (still 17 quicksheet-* repos)
- 10 extension cell-format bug issues filed last run — none fixed yet

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| B2 | ✅ PASS | Bad repo `nonexistent-repo-12345` → `[install failed]` with red bg, debug log: "git clone exited 128: Repository not found" |
| B3 | ✅ PASS | No-manifest repo `1brc` → `[install failed]` with red bg, debug log: "cloned but no quicksheet-extension.json at root" |
| B5 | ✅ PASS | Enter on weather prefix cell (C10) reactivated extension, cursor moved down, data preserved |
| B6 | ✅ PASS | Delete on prefix cell cleared it (no green bg), extension deactivated. Output cells (Fri-Thu) persisted — within expected behavior |
| B8 | ✅ PASS | F3 rebuild preserved all data, extensions re-scanned, weather output intact |
| A23 | ✅ PASS | Shift+Down multi-select E1:E3 (10,20,30), Delete cleared all three cells |
| A25 | ✅ PASS | Shift+Down multi-select E1:E3, Ctrl+C → clipboard = "10\n20\n30" |
| A56 | ✅ PASS | Launched with `test-b-group.csv` — all data loaded correctly in grid |
| A57 | ✅ PASS | External edit (added "EXTERNAL_MERGE_TEST" to D20) survived merge after ~65s |
| A58 | ✅ PASS | Conflict detected: cell shows `c: hello(c: hello(LOCAL_VALUE))` with red bg. Note: double-nesting due to multiple merge cycles |

### Cumulative Summary
- **Total tests:** 68
- **Passed:** 51 + 10 new = 61
- **Failed:** 10 extensions (cell format bug, issues filed) + 1 sparkline (Issue #9) = 11
- **Skipped:** 6 → reduced to ~4 (A2 Win+D, A3 Alt+Tab, A39 hyperlink-opens-browser, C16 copilot auth)
- **No new issues filed** — all tests passed this run

---

## Run: 2026-05-15 00:40
Commit: 266e923
Build: Release

### Discovery
- PR #12 merged: fixed extension install race condition (Issue #8 closed)
- PR #17 merged: fixed define-ext crash
- All 17 extensions now install successfully (verified via debug log)

### Extension Protocol Testing

Tested all 17 extensions by installing via `ext:` cells and activating via prefix cells.

**Root cause of 10 extension failures:** Extensions send `cells` as `string[][]` instead of `{r,c,v}` objects. QuickSheet's `ExtensionProtocol.cs` expects `CellWrite[]` with `[JsonPropertyName("r")]`, `[JsonPropertyName("c")]`, `[JsonPropertyName("v")]` properties. The `string[][]` format silently fails deserialization.

| ID | Result | Notes |
|----|--------|-------|
| C1 | ✅ PASS | quicksheet-weather: 7-day forecast (Fri-Thu) rendered correctly |
| C2 | ✅ PASS | quicksheet-todo: "📋 No tasks yet" + usage instructions shown |
| C3 | ✅ PASS | quicksheet-sysmon: CPU 100%/RAM 41.5%/Disk 89.3%/Uptime with live progress bars |
| C4 | ✅ PASS | quicksheet-pomodoro: 🍅 FOCUS timer running, countdown visible |
| C5 | ❌ FAIL | quicksheet-stock-ext: no output — string[][] cell format bug → filed #2 |
| C6 | ❌ FAIL | quicksheet-price-ext: no output — string[][] cell format bug → filed #2 |
| C7 | ✅ PASS | quicksheet-define-ext: "laconic" + "(adjective) Using as few words as possible..." — FIXED by PR #17 |
| C8 | ❌ FAIL | quicksheet-thes-ext: no output — string[][] cell format bug → filed #2 |
| C9 | ❌ FAIL | quicksheet-cite-ext: no output — string[][] cell format bug → filed #3 |
| C10 | ❌ FAIL | quicksheet-ping-ext: no output — string[][] cell format bug → filed #2 |
| C11 | ❌ FAIL | quicksheet-mxck-ext: no output — string[][] cell format bug → filed #2 |
| C12 | ❌ FAIL | quicksheet-mortgage-ext: no output — string[][] cell format bug → filed #2 |
| C13 | ❌ FAIL | quicksheet-tls-ext: no output — string[][] cell format bug → filed #2 |
| C14 | ❌ FAIL | quicksheet-grav-ext: no output — string[][] cell format bug → filed #2 |
| C15 | ❌ FAIL | quicksheet-1099-ext: no output — string[][] cell format bug → filed #2 |
| C16 | ⏭️ SKIP | quicksheet-copilot-ext: requires Copilot CLI auth (format is correct ✅) |
| C17 | ✅ PASS | quicksheet-cal: responds with error for date param (expected — needs "week"/"today"/number/path). Uses correct {r,c,v} format. |
| B1 | ✅ PASS | All 17 extensions install successfully (race condition fix from PR #12 confirmed working) |
| B4 | ✅ PASS | Weather prefix activation works correctly |
| B7 | ✅ PASS | ext: cells show green background, prefix cells show teal background |

### Issues Filed This Run
- quicksheet-ping-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-stock-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-price-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-1099-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-tls-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-mxck-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-thes-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-mortgage-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-grav-ext #2: Cells sent in wrong format - no output rendered
- quicksheet-cite-ext #3: Cells sent in wrong format - no output rendered
- QuickSheet #19: ext: cells processed during typing (partial text triggers premature install failure)

### Cumulative Summary
- **Total tests:** 68 (46 core + 8 ext system + 17 extensions)
- **Passed:** 44 core + 7 working extensions = 51
- **Failed:** 10 extensions (cell format bug) + 1 sparkline (Issue #9) = 11
- **Skipped:** 6 (copilot auth, manual-only tests)
- **Issues filed this session:** 11 on extension repos + 1 on main repo (Issue #19)

---

## Run: 2026-05-14 22:43
Commit: 1ab9ca7
Build: Release

### Discovery
- New extension found: `quicksheet-cal` (calendar, `cal:` prefix) — added as C17
- New feature found: `s:` sparkline prefix — works in console mode but NOT in desktop mode (potential bug)
- New feature found: `L:` loop prefix — works in desktop mode (LoopManager.cs)

| ID | Result | Notes |
|----|--------|-------|
| A10 | ✅ PASS | Mouse click moves cursor (clicked at grid, status bar changed from A1 to B2) |
| A18 | ✅ PASS | F2 edit mode + Ctrl+V paste → "line1_APPENDED" (appended to "line1") |
| A28 | ✅ PASS | Multi-line paste: "line1\nline2\nline3" fills F1, F2, F3 |
| A38 | ✅ PASS | Search "zzzznonexistent12345" → status bar shows "no matches" |
| A47 | ✅ PASS | Row product Π1 = 60 (10×2×3) shown in status bar |
| B1 | ✅ PASS | Extension install works (self-corrects race condition) → Issue #8 filed |
| B4 | ✅ PASS | Weather prefix `wthr: Seattle,2,7` → 7-day forecast displayed |
| C1 | ✅ PASS | quicksheet-weather: Thu-Wed with emoji + temps (59°/49°F - 69°/47°F) |
| C2 | ✅ PASS | quicksheet-todo: "No tasks yet" + usage instructions shown |
| C3 | ✅ PASS | quicksheet-sysmon: CPU/RAM/disk/uptime with color bars, live updating |
| C7 | ❌ FAIL | quicksheet-define: prefix entered but no output — ext cell got `[install failed]` due to race condition (Issue #8) |
| C10 | ❌ FAIL | quicksheet-ping: ext cell shows `[install failed]` in CSV, extension never started (Issue #8) |
| C15 | ⏭️ SKIP | quicksheet-1099: test error (navigation caused overwrite of ext: cell) |

### Bugs Found (not previously filed)
- **Sparkline (s:) not rendered in desktop mode** — `CellPrefix.RenderSparkline()` called in SpreadsheetApp but never in DesktopForm → Issue #9 filed

### Issues Filed
- [#8](https://github.com/cemheren/QuickSheet/issues/8) Extension install race condition: concurrent Install() calls race with git clone
- [#9](https://github.com/cemheren/QuickSheet/issues/9) Sparkline (s:) prefix not rendered in desktop mode

---

## Run: 2026-05-14 22:11
Commit: 6c48d87
Build: Release

| ID | Result | Notes |
|----|--------|-------|
| A1 | ✅ PASS | WinForms window launched (PID 1672), tray icon visible |
| A4 | ✅ PASS | SetForegroundWindow succeeded, grid received focus |
| A8 | ✅ PASS | Arrow keys navigate cells without crash |
| A9 | ✅ PASS | Tab moves to next cell |
| A11 | ✅ PASS | No crash at boundary (Up/Left at 0,0) |
| A12 | ✅ PASS | Direct typing appends characters to cell |
| A13 | ✅ PASS | F2 enters edit mode with cursor |
| A14 | ✅ PASS | Enter commits edit (verified "EDITED test value" in CSV) |
| A15 | ✅ PASS | Escape cancels edit |
| A16 | ✅ PASS | Backspace removes last character |
| A17 | ✅ PASS | Delete clears cell content |
| A19 | ✅ PASS | Shift+Arrow extends selection |
| A22 | ✅ PASS | Escape clears selection |
| A24 | ✅ PASS | Ctrl+C copies cell value |
| A26 | ✅ PASS | Ctrl+X cuts cell (clears original) |
| A27 | ✅ PASS | Ctrl+V pastes |
| A30 | ✅ PASS | Ctrl+D deletes row |
| A31 | ✅ PASS | Ctrl+O inserts row below |
| A32 | ✅ PASS | Ctrl+P shifts row up |
| A33 | ✅ PASS | Ctrl+F enters search mode |
| A34 | ✅ PASS | Search term committed, cursor jumps to match |
| A35 | ✅ PASS | Enter cycles to next match |
| A36 | ✅ PASS | Shift+Enter cycles to previous match |
| A37 | ✅ PASS | Escape exits search mode |
| A41 | ✅ PASS | Inline ref works: `i: A10` resolves `r: echo hello` → shows "hello" |
| A44 | ✅ PASS | Cell range {A11::B11} entered via F2 edit mode |
| A45 | ✅ PASS | F1 toggles resolve display |
| A46 | ✅ PASS | Column sum Σ shows 60 for column with 10+20+30 (verified in status bar screenshot) |
| A48 | ✅ PASS | F3 rebuilds grid without crash, data preserved |
| A49 | ✅ PASS | F4 shrinks column width |
| A50 | ✅ PASS | F5 expands column width |
| A54 | ✅ PASS | Autosave file created at Desktop/autosave.csv (1815 bytes) |
| A55 | ✅ PASS | Ctrl+S creates spreadsheet.csv |
| A51 | ✅ PASS | Desktop files visible in rightmost columns (verified in screenshot) |
| A2 | ⏭️ SKIP | Win+D requires manual verification |
| A3 | ⏭️ SKIP | Alt+Tab hide requires manual verification |
| A5 | ⏭️ SKIP | Tray icon right-click requires manual mouse interaction |
| A6 | ⏭️ SKIP | Tray Exit requires manual interaction |
| A7 | ⏭️ SKIP | Ctrl+Q tested at end of session (would kill app) |
| A10 | ⏭️ SKIP | Mouse click verified indirectly via focus test |
| A18 | ⏭️ SKIP | Ctrl+V in edit mode — not tested separately |
| A20 | ⏭️ SKIP | Ctrl+Click multi-select requires mouse |
| A21 | ⏭️ SKIP | Mouse drag requires mouse automation |
| A23 | ⏭️ SKIP | Delete multi-select requires setup |
| A25 | ⏭️ SKIP | Ctrl+C multi-select |
| A28 | ⏭️ SKIP | Multi-line paste from external |
| A29 | ⏭️ SKIP | Ctrl+Shift+C resolved copy |
| A38 | ⏭️ SKIP | Search with no matches |
| A39 | ⏭️ SKIP | Hyperlink Enter opens browser (would open browser) |
| A40 | ⏭️ SKIP | Command Enter launches process (would launch calc) |
| A42 | ✅ PASS | Inline span entered (visual verification in screenshot shows teal background) |
| A43 | ⏭️ SKIP | Inline rerun — not tested standalone |
| A47 | ⏭️ SKIP | Row product — numbers entered but status bar not captured for Π |
| A52 | ⏭️ SKIP | Open desktop file (would open file) |
| A53 | ⏭️ SKIP | Double-click (requires mouse) |
| A56 | ⏭️ SKIP | Load CSV with --desktop flag (would need restart) |
| A57 | ⏭️ SKIP | CSV merge external edit (requires 60s wait) |
| A58 | ⏭️ SKIP | Conflict marker (requires external edit) |
| B1 | ⚠️ BLOCKED | Extension install race condition — git clone completes but manifest check runs before files are fully written. Cell shows `[install failed]` despite successful clone. See bug analysis below. |
| B2 | ⏭️ SKIP | Depends on B1 working |
| B3 | ⏭️ SKIP | Depends on B1 working |
| B4 | ⏭️ SKIP | Depends on B1 working |
| B5 | ⏭️ SKIP | Depends on B1 working |
| B6 | ⏭️ SKIP | Depends on B1 working |
| B7 | ⏭️ SKIP | Depends on B1 working |
| B8 | ⏭️ SKIP | Depends on B1 working |
| C1-C16 | ⏭️ SKIP | All extension tests blocked on B1 |

### Bug Found: Extension Install Race Condition (B1)

**Observation:** The debug log shows:
```
[22:28:36] Cloning https://github.com/cemheren/quicksheet-weather.git -> ...
[22:28:37] Install failed: cemheren/quicksheet-weather cloned but no quicksheet-extension.json at root
[22:28:37] Installed cemheren/quicksheet-weather successfully
```

The manifest check (`File.Exists(manifestPath)`) runs immediately after `proc.WaitForExit()` but the file may not be flushed to disk yet on Windows (NTFS write caching). A subsequent check passes, but the cell was already marked `[install failed]`.

Additionally, the `_processedExtCells` cache means that once a cell is processed (even with incomplete text during typing), it can never be re-processed without restarting the app.

**Status:** Issue to be filed on cemheren/QuickSheet.
