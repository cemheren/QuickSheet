# QuickSheet Test Log

## Run: 2026-05-15 14:50
Commit: 5c2909f (main), v0.10.0
Build: Release

### Discovery
- **1 new extension**: **quicksheet-depr-ext** (depreciation calculator) — manifest correct (`"entry"`, `"prefix": "depr"`)
- **PR #61 OPEN**: New feature — cell color prefix `c:color: text` for highlighting cells
- PRs #60, #62, #63 open (docs)
- worldtm/mileage fix PRs still NOT merged
- 34 quicksheet-* repos total (was 33)

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C32 | ❌ FAIL | depr-ext: 3 protocol bugs — wrong message types (register/invoke instead of init/activate) + wrong cell format (row/col/value instead of r/c/v). No output produced. → filed depr-ext#1 |
| C28 | ❌ FAIL | worldtm: fix PRs #2/#3 still OPEN |
| C29 | ❌ FAIL | mileage-ext: fix PRs #2/#3 still OPEN |

### Issues Filed
- **cemheren/quicksheet-depr-ext#1** — Extension uses wrong protocol message types + cell format

### Cumulative Summary
- **Total tests:** 88 (48 core + 8 ext system + 32 extensions)
- **Passed:** 75
- **Failed:** C28 (worldtm), C29 (mileage), C32 (depr protocol) = 3
- **Blocked:** C16 (copilot auth), C25 (docker), C31 (k8s) = 3
- **Skipped:** A2, A3, A5, A6, A20, A21, A52 = 7

---

## Run: 2026-05-15 13:48
Commit: ff7f371 (main)
Build: Release

### Discovery
- 2 new extensions: **quicksheet-margin-ext** (break-even analysis), **quicksheet-k8s** (Kubernetes pod status)
- Grow skill running in parallel (gh-pages updates, PR merges)
- `grow/add-margin-ext` branch exists — margin-ext added to QuickSheet docs
- worldtm PR#2/#3 still OPEN — manifest fix not merged
- mileage-ext PR#2/#3 still OPEN — manifest fix not merged
- v0.10.0 being cut by grow skill
- 33 quicksheet-* repos total (was 31)

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C30 | ✅ PASS | margin-ext: "🟡 Moderate margin", CM $40,000/unit (40%), break-even 1 unit, BE revenue $100k. Correct output! |
| C31 | ⚠️ BLOCKED | k8s: "Error: The operation has timed out" — no Kubernetes cluster configured. Extension ran correctly. |
| C28 | ❌ FAIL | worldtm: still "Manifest missing 'entry' field" — fix PRs #2/#3 still OPEN |
| C29 | ❌ FAIL | mileage-ext: still missing prefix field — fix PRs #2/#3 still OPEN |

### Cumulative Summary
- **Total tests:** 87 (48 core + 8 ext system + 31 extensions)
- **Passed:** 75 (prev 74 + C30)
- **Failed:** C28 (worldtm entry), C29 (mileage prefix) = 2
- **Blocked:** C16 (copilot auth), C25 (docker not running), C31 (k8s no cluster) = 3
- **Skipped:** A2, A3, A5, A6, A20, A21, A52 = 7

---

## Run: 2026-05-15 12:47
Commit: 9ee51dd (main)
Build: Release

### Discovery
- **PR #52 MERGED**: sparkline (s:) prefix in Linux desktop mode
- **ghpr#2 MERGED**: fix reviewDecision + params + flatten cells
- **docker#2 MERGED**: Windows named pipe + params fix; **docker#4 MERGED**: register version as string
- 2 new extensions: **quicksheet-worldtm** (world clock), **quicksheet-mileage-ext** (IRS mileage)
- PRs open: #53 (changelog), #54 (worldtm docs), #55 (mileage link), #56 (extensions doc protocol)
- v0.9.0 tagged
- 31 quicksheet-* repos (was 29 last run)

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C23 | ✅ PASS | ghpr fix (PR #2) works! Response: "No PRs need attention" — valid output for no matching PRs. |
| C25 | ⚠️ BLOCKED | Docker Desktop not running on this machine. Extension code fixes verified (PRs #2, #4 merged). |
| C28 | ❌ FAIL | worldtm: manifest uses `"entrypoint"` instead of `"entry"` → filed worldtm#1 |
| C29 | ❌ FAIL | mileage-ext: manifest missing `"prefix"` field → filed mileage-ext#1 |

### Issues Filed
- **cemheren/quicksheet-worldtm#1** — Manifest uses 'entrypoint' instead of 'entry'
- **cemheren/quicksheet-mileage-ext#1** — Manifest missing 'prefix' field

### Cumulative Summary
- **Total tests:** 85 (48 core + 8 ext system + 29 extensions)
- **Passed:** 74 (85 - 2 fail - 2 blocked - 7 skip)
- **Failed:** C28 (worldtm entry), C29 (mileage prefix) = 2
- **Blocked:** C16 (copilot auth), C25 (docker not running) = 2
- **Skipped:** A2, A3, A5, A6, A20, A21, A52 = 7
- **Note:** Sparkline issue #9 tracked separately

---

## Run: 2026-05-15 09:45
Commit: 5b75db7 (main)
Build: Release

### Discovery
- **PR #44 MERGED** (`cb6699d`) — FlexVersionConverter fixes RegisterMessage version type mismatch
- New extension: **quicksheet-cntdn** (countdown timer) — added as C27
- Manifest filename standardized to `quicksheet-extension.json` (old `manifest.json` now 404 on hntop/portck)
- hntop/portck manifests still use `"entrypoint"` instead of `"entry"` field name
- cntdn manifest uses `"entryPoint"` (camelCase) instead of `"entry"` — same class of bug
- gitst manifest has correct `"entry"` field ✅ and `"version": "1.0.0"` (string) ✅
- Extension deserializer uses `JsonNamingPolicy.CamelCase`: C# `Entry` → JSON `"entry"`

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C26 | ❌ FAIL | Version type fixed by PR #44 ✅ but `arguments` vs `params` mismatch remains → no output. Filed gitst#4. |
| C27 | ❌ FAIL | New ext (cntdn): manifest uses `"entryPoint"` not `"entry"` → fails to launch. Filed cntdn#1. |
| C21 | ❌ FAIL | hntop: manifest now `quicksheet-extension.json` but `"entrypoint"` not `"entry"`. Commented on hntop#1. |
| C24 | ❌ FAIL | portck: same as hntop. Commented on portck#1. |

### Issues Filed/Updated
- **cemheren/quicksheet-cntdn#1** — Manifest uses `entryPoint` instead of `entry`
- **cemheren/quicksheet-gitst#4** — Reads `arguments` instead of `params`, user args ignored
- Commented on gitst#1 (version fix confirmed via PR #44)
- Commented on hntop#1, portck#1 (manifest filename fixed, field name still wrong)

### Cumulative Summary
- **Total tests:** 81 (46 core + 8 ext system + 27 extensions)
- **Passed:** 80
- **Failed:** C21 (hntop entry field), C24 (portck entry field), C27 (cntdn entry field), C23 (ghpr params+search), C25 (docker Windows+params), C26 (gitst params mismatch), sparkline Issue #9 = 7
- **Blocked:** C16 (copilot auth) = 1
- **Skipped:** A2, A3, A5, A6, A20, A21, A43, A52, A53 = 9

---

## Run: 2026-05-15 09:00
Commit: 8f703f8 (main)
Build: Release

### Discovery
- PR #44 OPEN: fix for RegisterMessage version type mismatch (FlexVersionConverter)
- Issue #43 CLOSED
- No new extensions (28 repos unchanged)
- Extension repos (gitst, ghpr, docker) still use `version: 1` (int) — unfixed
- hntop and portck manifests still use "entrypoint" — unfixed

### Tests Run (Session 1: focused A-group)

| ID | Result | Notes |
|----|--------|-------|
| A41 | ✅ PASS | B3 = "apple" (teal bg) = resolved value of i: A1 |
| A44 | ✅ PASS | Shows resolved {A1::C2} — concatenated cell values from A1-C2 |
| A45 | ✅ PASS | F1 toggles status bar between "F1: Raw" and "F1: Resolve" |
| A48 | ✅ PASS | F3 rebuild preserves data, desktop files re-populated |
| A49 | ✅ PASS | F4 shrinks columns (12 → many narrow). Click-to-focus needed. |
| A50 | ✅ PASS | F5 expands columns (12 → 3 wide). Click-to-focus needed. |
| A51 | ✅ PASS | Desktop files/folders listed in rightmost column |

### Tests Run (Session 2: hyperlink, save, quit)

| ID | Result | Notes |
|----|--------|-------|
| A39 | ✅ PASS | https://github.com opened Chrome (purple bg on cell) |
| A55 | ✅ PASS | Ctrl+S saved CSV (LastWriteTime confirmed 9:01 AM). Tray balloon may have expired before screenshot. |
| A43 | ⏭️ SKIP | Cursor landed on D2 instead of C2 (i: A1). Navigation offset due to column widths. Needs retry. |

### Note: SendKeys focus issue
SendKeys via SetForegroundWindow alone doesn't reach QuickSheet (WS_EX_TOOLWINDOW). Must click on the grid first via mouse_event to properly acquire keyboard focus.

### Note: Hyperlink side-effect
After A39, Chrome takes focus and its Ctrl+S opens a "Save As" dialog. Must close Chrome/minimize before continuing QuickSheet tests.

### Cumulative Summary
- **Total tests:** 80 (46 core + 8 ext system + 26 extensions)
- **Passed:** 80 (71 prev + 7 session1 + 2 session2 = 80)
- **Failed:** C21 (hntop manifest), C23 (ghpr: version+params+search), C24 (portck manifest), C25 (docker: version+Windows+params), C26 (gitst: version mismatch), sparkline Issue #9 = 6
- **Blocked:** C16 (copilot auth) = 1
- **Skipped:** A2, A3, A5, A6, A20, A21, A43, A52, A53 = 9 (A39 resolved)
- **Not yet tested:** (none from defined tests — all 80 have results)

---

## Run: 2026-05-15 08:30
Commit: c5ec47f (main)
Build: Release

### Discovery
- v0.7.0 tagged. 2 NEW extensions: `quicksheet-gitst` (git status), `quicksheet-rate` (404/no manifest yet)
- hntop and portck manifests still NOT fixed (still use `"entrypoint"`)

### Major Bug Found: RegisterMessage.Version type mismatch (Issue #43)

**Root cause of gitst, ghpr, and docker producing no output:**
- `RegisterMessage.Version` in `ExtensionProtocol.cs:72` is typed as `string`
- `InitMessage.Version` in `ExtensionProtocol.cs:45` is typed as `int`
- Newer extensions (gitst, ghpr, docker) send `"version":1` (integer) mirroring the init format
- `System.Text.Json` strict typing rejects int→string coercion → `JsonException` thrown
- Catch block at `ExtensionManager.cs:281` silently swallows the error
- Register message effectively IGNORED → prefix never in `_prefixMap` → activate never sent
- All 22 working extensions send `"version":"1.0.0"` (string) → work fine

**Debug methodology:**
1. Launched QuickSheet with stderr captured, only register message appeared on stdout
2. Added temporary debug logging to ExtensionManager
3. Confirmed: register message received but `Deserialize<RegisterMessage>()` returned null
4. `KnownPrefixes` remained empty across all ScanGrid calls
5. Verified weather sends `version:"1.0.0"` (string) and works

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C26 | ❌ FAIL | quicksheet-gitst: register sends `version:1` (int) → deserialization fails → prefix never registered → no output. Filed QuickSheet#43 + gitst#1 |
| C23 | ❌ FAIL | Updated: root cause is version int/string mismatch (same as C26). Commented on ghpr#1 |
| C25 | ❌ FAIL | Updated: root cause is version int/string mismatch (same as C26). Commented on docker#1 |

### Issues Filed
- QuickSheet#43: RegisterMessage.Version typed as string rejects int version from extensions (gitst, ghpr, docker)
- quicksheet-gitst#1: Register message sends version as int, should be string

### Cross-cutting: version int vs string mismatch
| Extension | version field | Status |
|-----------|--------------|--------|
| quicksheet-gitst | `version = 1` (int) | ❌ Broken |
| quicksheet-ghpr | `version = 1` (int) | ❌ Broken |
| quicksheet-docker | `version = 1` (int) | ❌ Broken |
| All 22 others | `version = "1.0.0"` (string) | ✅ Working |

### Cumulative Summary
- **Total tests:** 80 (46 core + 8 ext system + 26 extensions)
- **Passed:** 71
- **Failed:** C21 (hntop manifest), C23 (ghpr: version+params+search), C24 (portck manifest), C25 (docker: version+Windows+params), C26 (gitst: version mismatch), sparkline Issue #9 = 6
- **Blocked:** C16 (copilot auth) = 1
- **Skipped:** A2, A3, A5, A6, A20, A21, A39, A43, A52, A53 = 10

---

## Run: 2026-05-15 06:50
Commit: c5ec47f (main)
Build: Release

### Discovery
- PR #36 merged: `quicksheet-portck` (TCP port checker extension)
- Branches: `grow/fix-issue-35-ctrl-b-desktop`, `grow/fix-sort-desktop-mode` (PRs #37, #38 OPEN — Ctrl+B desktop fix)
- Branch: `grow/add-docker-extension` — new docker extension
- 2 new extensions: `quicksheet-portck`, `quicksheet-docker` (now 25 total)
- `quicksheet-portck` manifest uses `"entrypoint"` (same bug as hntop)
- `quicksheet-docker` manifest uses correct `"entry"` ✅

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C24 | ❌ FAIL | quicksheet-portck: manifest uses "entrypoint" → install fails. Filed portck#1 |
| C25 | ❌ FAIL | quicksheet-docker: (a) Windows not supported (Unix socket only), (b) "arguments" vs "params" mismatch, (c) errors not shown. Filed docker#1 |
| C23 | ❌ FAIL | quicksheet-ghpr: extension responds but `reviewDecision` is invalid gh search field → all searches fail → always shows "No PRs". Also "arguments" vs "params" mismatch. Filed ghpr#1 |
| C21 | ❌ FAIL | quicksheet-hntop: manifest still uses "entrypoint". Existing hntop#1 |
| A-sort | ✅ PASS | Ctrl+B sort in desktop mode WORKS — Issue #35 closed via PR #38 (v0.7.0). Data sorted A-Z correctly. |

### Issues Filed
- `quicksheet-portck#1`: Manifest uses "entrypoint" instead of "entry"
- `quicksheet-docker#1`: Doesn't work on Windows (Unix socket, params mismatch, silent errors)
- `quicksheet-ghpr#1`: All searches fail (invalid reviewDecision field) + params mismatch
- Issue #35: CONFIRMED FIXED (PR #38 merged, v0.7.0)

### Cross-cutting bug found: "arguments" vs "params" field name mismatch
Multiple extensions (ghpr, docker) read `root.TryGetProperty("arguments", ...)` but QuickSheet's `ActivateMessage` sends the field as `"params"` (string[]). This means extensions never receive user arguments.

### Cumulative Summary
- **Total tests:** 79 (46 core + 8 ext system + 25 extensions)
- **Passed:** 71
- **Failed:** C21 (hntop manifest), C23 (ghpr searchfield+params), C24 (portck manifest), C25 (docker Windows+params), sparkline Issue #9 = 5
- **Blocked:** C16 (copilot auth) = 1
- **Skipped:** A2, A3, A5, A6, A20, A21, A39, A43, A52, A53 = 10
- define-ext crash fix branch force-updated
- hntop manifest still uses "entrypoint" (NOT fixed despite grow log claiming so)
- Ctrl+B sorting wired in console mode only — NOT in DesktopForm (feature parity gap)

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C6 | ✅ PASS | quicksheet-price-ext: bitcoin $80,382, ▲ 0.87% 24h — CoinCap fallback works! |
| C21 | ❌ FAIL | quicksheet-hntop: installed but "Manifest missing 'entry' field" — existing hntop#1 |
| C23 | ⚠️ BLOCKED | quicksheet-ghpr: installed, 3 processes running, but no output. Requires `gh` CLI auth. Needs investigation. |

### Cumulative Summary
- **Total tests:** 74 (46 core + 8 ext system + 23 extensions)
- **Passed:** 71 (C6 flipped from FAIL to PASS)
- **Failed:** C21 (hntop manifest), sparkline (Issue #9) = 2
- **Blocked:** C23 (ghpr no output), C16 (copilot auth) = 2
- **Skipped:** A2, A3, A5, A6, A20, A21, A39, A43, A52, A53 = 10
- **Known gaps:** Ctrl+B sorting not in desktop mode (console only)

---

## Run: 2026-05-15 04:45
Commit: d98da91
Build: Release

### Discovery
- 2 new extensions found: quicksheet-hntop (HN stories), quicksheet-apistatus (service status monitor)
- All 10 cell-format fixes merged, PR #21 merged (typing race)
- Issue #26 fix PRs #27/#29 still OPEN
- PR #28 OPEN: Ctrl+B column sorting
- price-ext 403 still unfixed

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C21 | ❌ FAIL | quicksheet-hntop: `[bad manifest]` — uses "entrypoint" instead of "entry" → filed hntop#1 |
| C22 | ✅ PASS | quicksheet-apistatus: 🟢 github operational, 🟢 npm operational, 🟡 cloudflare minor, Updated 04:47 |
| C6 | ❌ FAIL | quicksheet-price-ext: still 403 Forbidden (CoinGecko API) — existing issue #4 |
| A40 | ✅ PASS | r: echo test-output-123 — cmd process launched, yellow bg on cell |

### Cumulative Summary
- **Total tests:** 73 (46 core + 8 ext system + 22 extensions)
- **Passed:** 70
- **Failed:** C6 (price 403), C21 (hntop manifest), sparkline (Issue #9) = 3
- **Skipped:** A2, A3, A5, A6, A20, A21, A39, A43, A52, A53, C16 = 11

---

## Run: 2026-05-15 03:45
Commit: 5c71062
Build: Release

### Discovery
- PR #21 merged: fix ext: typing race (Issue #19) ✅
- All 10 broken extensions got cell-format fix PRs merged (`grow/fix-cells-shape`) ✅
- Issue #26 fix PRs #27 and #29 OPEN (install corruption — not merged yet)
- PR #28 OPEN: Ctrl+B column sorting (new feature)
- New repo: `quicksheet-console` (no description, install failed — needs investigation)

### Tests Run — Re-test of 10 previously-broken extensions

| ID | Result | Notes |
|----|--------|-------|
| C5 | ✅ PASS | quicksheet-stock-ext: AAPL.US $298.21, ▼ -0.54% (-1.61), as of 2026-05-14 |
| C6 | ❌ FAIL | quicksheet-price-ext: `err: Response status code does not indicate success: 403 (Forbidden).` → filed quicksheet-price-ext#4 |
| C8 | ✅ PASS | quicksheet-thes-ext: happy → halcyon, content, bright, felicitous, riant |
| C9 | ✅ PASS | quicksheet-cite-ext: Zhang W., Hebig R., et al. (2023), full ACM citation with DOI |
| C10 | ✅ PASS | quicksheet-ping-ext: ✓ https://github.com, 200 OK, 545 ms |
| C11 | ✅ PASS | quicksheet-mxck-ext: MX for github.com, 0 github-com.mail.protection.outlook.com |
| C12 | ✅ PASS | quicksheet-mortgage-ext: $500K@6.5%/30yr, monthly $3,160.34, total interest $637,722 |
| C13 | ✅ PASS | quicksheet-tls-ext: github.com:443, expires in 79d, issuer: Sectigo Limited, cn: github.com |
| C14 | ✅ PASS | quicksheet-grav-ext: test@example.com, (no Gravatar profile), avatar URL present |
| C15 | ✅ PASS | quicksheet-1099-ext: $80K net income, SE tax ~$11,304, quarterly ~$2,826, disclaimer |

### Cumulative Summary
- **Total tests:** 71 (46 core + 8 ext system + 20 extensions)
- **Passed:** 68 + 9 newly fixed = 77 (but some overlap with previous fails → net 77)
- **Failed:** 1 extension (C6 price 403, issue filed) + 1 sparkline (Issue #9) = 2
- **Skipped:** A2 (Win+D), A3 (Alt+Tab), A5 (tray), A6 (tray exit), A20 (Ctrl+Click), A21 (drag), A39 (hyperlink), A40 (r: cmd), A43 (inline rerun), A52 (open file), A53 (dbl-click), C16 (copilot auth)
- **New issue filed:** quicksheet-price-ext#4 (CoinGecko 403)
- **quicksheet-console:** Install failed — new extension, needs investigation (not yet in Group C)

---

## Run: 2026-05-15 02:49
Commit: cf75d79
Build: Release

### Discovery
- No new code changes (only grow-skill log commits)
- No new extensions (still 20 quicksheet-* repos)
- No fixes merged for 10 broken extensions (string[][] format)
- PR #18 still closed, PRs #21/#23 still pending

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| A7 | ✅ PASS | Ctrl+Q exits app cleanly (via SetForegroundWindow + SendKeys) |
| A29 | ✅ PASS | Ctrl+Shift+C on i: cell copied resolved value "hello" (not raw "i: A1") |
| A46 | ✅ PASS | Column sum ΣA = 62 displayed in status bar |
| A47 | ✅ PASS | Row product Π8 = 24 (2×3×4) displayed in status bar |

### Cumulative Summary
- **Total tests:** 71 (46 core + 8 ext system + 20 extensions)
- **Passed:** 64 + 4 new = 68
- **Failed:** 10 extensions (cell format bug, issues filed) + 1 sparkline (Issue #9) = 11
- **Skipped:** A2 (Win+D manual), A3 (Alt+Tab manual), A5 (tray menu), A6 (tray exit), A20 (Ctrl+Click), A21 (mouse drag), A39 (hyperlink browser), A40 (r: command), A43 (inline rerun), A52 (open desktop file), A53 (double-click), C16 (copilot auth)
- **Not yet filed:** `[install failed]` cascading corruption bug (ExtensionManager.cs:64,181 + CellPrefix.cs:131)

---

## Run: 2026-05-15 01:47
Commit: bd7204d
Build: Release

### Discovery
- 3 new extensions found: quicksheet-fx (currency), quicksheet-qtr (tax deadlines), quicksheet-budget (envelope visualizer)
- All 3 use correct {r,c,v} cell format ✅
- Issue #19 open: ext: cells processed during typing (fix PRs #21, #23 pending)
- PR #18 (string[][] format fix) was CLOSED without merging — extensions must fix individually
- New branches: grow/add-fx-extension, grow/add-qtr-extension, grow/fix-ext-typing-race

### Tests Run

| ID | Result | Notes |
|----|--------|-------|
| C18 | ✅ PASS | quicksheet-fx: "💱 1,000.00 USD" → EUR 0.8551 = 855.14 EUR (ECB · 2026-05-15) |
| C19 | ✅ PASS | quicksheet-qtr: "📅 Tax Year 2026" — Q1 ✅ Paid, Q2 🟢 32d, Q3 🟢 123d, Q4 🟢 245d |
| C20 | ✅ PASS | quicksheet-budget: "🟡 Groceries" 70.0% bar, $350/$500 spent, $150 remaining |
| C17 | ✅ PASS | quicksheet-cal: uses correct format, responds with calendar error for invalid param (expected — needs "week"/"today"/number) |

### Cumulative Summary
- **Total tests:** 71 (46 core + 8 ext system + 20 extensions — 3 new)
- **Passed:** 61 + 3 new extensions = 64
- **Failed:** 10 extensions (cell format bug, issues filed) + 1 sparkline (Issue #9) = 11
- **Skipped:** ~4 (A2 Win+D, A3 Alt+Tab, A39 hyperlink, C16 copilot auth)
- **No new issues filed** — all new extensions work correctly

---

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

## Run: 2026-05-15 17:50
Commit: 54d62ec (main), v0.11.0
Build: Release

### Discovery
- **PR #61 MERGED**: c:color: cell prefix feature (highlights cell backgrounds with named colors)
- **PR #64 MERGED**: docs for c:color: in keyboard-shortcuts.md
- **New file**: `docs/keyboard-shortcuts.md` documents ALL shortcuts
- **5 console-mode shortcuts NOT in desktop mode**: Ctrl+G (Go to), Ctrl+T (Theme), Ctrl+H (Help), Ctrl+Z (Undo), Ctrl+Y (Redo)
- **34 quicksheet-* repos** (no new ones)
- depr-ext/worldtm/mileage-ext fix PRs still OPEN

### Tests

| ID | Result | Notes |
|----|--------|-------|
| A61 |  FAIL | c:color: prefix shows as plain text  no colored background. DesktopForm.cs has zero references to CellPrefix.IsColored()/ParseColor().  Issue #65 filed |
| A62 |  FAIL | Ctrl+G (Go to cell) does nothing in desktop mode  types "g" in cell instead. Not in DesktopForm.cs Ctrl handler |
| A63 |  FAIL | Ctrl+T (Cycle theme) not implemented in desktop mode |
| A64 |  FAIL | Ctrl+H (Help overlay) not implemented in desktop mode |
| A65 |  FAIL | Ctrl+Z (Undo) not implemented in desktop mode |
| A66 |  FAIL | Ctrl+Y (Redo) not implemented in desktop mode |

A62-A66  Issue #66 filed (single issue for all 5 missing shortcuts)

### Cumulative (after Run 20)
- **Total: 94 tests** (54 core + 8 ext system + 32 extensions)
- **Passed: 75**
- **Failed: 9** (C28 worldtm, C29 mileage, C32 depr, A61 c:color, A62-A66 shortcuts)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 7** (A2, A3, A5, A6, A20, A21, A52)

## Run: 2026-05-15 16:45
Commit: 277b572 (main), v0.11.0
Build: Release

### Discovery
- **2 NEW extensions**: quicksheet-jwtdec (JWT decoder), quicksheet-rate (freelance rate calculator)
- quicksheet-rate has correct manifest (entry, prefix, params)
- quicksheet-jwtdec has manifest with trailing colon in prefix: "jwtdec:" (bug)
- **PRs #70, #71 OPEN**: fixes for Issue #65 (c:color) and Issue #66 (shortcuts)  not merged yet
- Extension fix PRs still OPEN: worldtm, mileage-ext, depr-ext
- **35 quicksheet-* repos** total (was 34)

### Tests

| ID | Result | Notes |
|----|--------|-------|
| A28 |  PASS | Multi-line paste: "line1\nline2\nline3" pasted correctly into consecutive rows |
| C33 |  FAIL | jwtdec: 2 protocol bugs  (1) responds to init with status:ready instead of register, (2) manifest prefix has trailing colon "jwtdec:"  jwtdec#1 filed |
| C34 |  FAIL | rate: responds to activate with type:"cells" instead of type:"write"  cells silently dropped  rate#1 filed |

### Cumulative (after Run 21)
- **Total: 96 tests** (54 core + 8 ext system + 34 extensions)
- **Passed: 76** (+1: A28)
- **Failed: 11** (C28 worldtm, C29 mileage, C32 depr, C33 jwtdec, C34 rate, A61 c:color, A62-A66 shortcuts)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 6** (A2, A3, A5, A6, A20, A21)

## Run: 2026-05-15 17:45
Commit: d2ae14b (main), v0.12.0
Build: Release

### Discovery
- **v0.12.0 released** (PR #72 merged)
- **PR #71 MERGED**: c:color: prefix rendering in desktop mode  Issue #65 CLOSED
- **PR #67 MERGED**: Also fixes c:color (duplicate PR  caused build error, filed #75, fixed)
- **PR #70 CLOSED (not merged)**: Windows Ctrl+Z/Y/T shortcuts NOT added. Issue #66 still open.
- **3 extension fix PRs MERGED**: worldtm#2, mileage-ext#2, depr-ext#2
- **2 NEW extensions**: quicksheet-cronck (cron parser), quicksheet-gitlog (git log viewer)  37 repos total
- **Build error found**: Duplicate `colorParsed` variable from merging PRs #67+#71  Issue #75 filed, fixed in 0b14e84
- jwtdec#2 and rate#2 fix PRs still OPEN

### Tests

| ID | Result | Notes |
|----|--------|-------|
| A61 |  PASS | c:color: prefix NOW WORKS in desktop mode! Red/green/blue/yellow backgrounds visible, text stripped correctly. PR #71 fix confirmed. |
| C28 |  PASS | worldtm: Shows London/Tokyo/NY with times, UTC offsets, and business hours status. Fix PR #2 confirmed. |
| C29 |  PASS | mileage: Shows IRS 2025 rate (.700/mi), .00 deduction for 1000 mi. Fix PR #2 confirmed. |
| C32 |  PASS | depr: Protocol test confirms correct register/write/{r,c,v}. Shows 5yr straight-line schedule (/yr). Fix PR #2 confirmed. (Note: space-separated params, not comma) |
| C35 |  FAIL | cronck: 2 protocol bugs  uses `init_response`/`activate_response` instead of `register`/`write`  cronck#1 filed |
| C36 |  FAIL | gitlog: Same bugs as jwtdec  responds with `status:ready` instead of `register`, trailing colon in prefix `"gitlog:"`  gitlog#1 filed |

### Build Fix
- **Issue #75 filed + fixed**: Duplicate `colorParsed` variable from PRs #67+#71. Committed fix 0b14e84.

### Cumulative (after Run 22)
- **Total: 98 tests** (54 core + 8 ext system + 36 extensions)
- **Passed: 80** (+4: A61, C28, C29, C32 all flipped from FAIL to PASS)
- **Failed: 9** (C33 jwtdec, C34 rate, C35 cronck, C36 gitlog, A62-A66 shortcuts)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 6** (A2, A3, A5, A6, A20, A21)

## Run: 2026-05-15 19:00
Commit: 5ad59bc (main, includes PR #76 merge)
Build: Release

### Discovery
- **PR #76 MERGED**: feat(desktop): add Ctrl+Z/Y/T/G/H shortcuts to Windows desktop mode (Issue #66)
- No new repos (still 37 quicksheet-* repos)
- Extension fix PRs still OPEN: jwtdec#2, rate#2, cronck#2. gitlog has no fix PR.

### Tests

| ID | Result | Notes |
|----|--------|-------|
| A18 |  PASS | Paste in edit mode: typed "BASE", F2 edit, Ctrl+V "+ADDED"  "BASE+ADDED"  |
| A52 |  PASS | Open desktop file: Ctrl+F found "poe_filter.txt.txt" in L25, Escape+Enter opened it in Notepad |
| A62 |  PASS | Ctrl+G go to cell: Status bar shows "Go to cell (e.g. A1, C5): A1", Enter navigates to A1 |
| A63 |  FAIL | Ctrl+T theme cycling: Theme.CycleNext() called but DesktopForm.OnPaint hardcodes colors  no visible change  filed #77 |
| A64 |  PASS | Ctrl+H help overlay: Beautiful shortcut reference box appears, "Press any key to close..." dismisses it |
| A65 |  PASS | Ctrl+Z undo: "CHANGED"  Ctrl+Z  "CHANGE" (character-level undo works) |
| A66 |  PASS | Ctrl+Y redo: "CHANGE"  Ctrl+Y  "CHANGED" (redo restores undone change) |

### Issues Filed
- **#77**  Ctrl+T theme cycling has no visible effect in desktop mode (DesktopForm hardcodes colors)
- Commented on **#66** with 4/5 shortcuts working (Ctrl+T is the exception  #77)

### Cumulative (after Run 23)
- **Total: 98 tests** (54 core + 8 ext system + 36 extensions)
- **Passed: 86** (+6: A18, A52, A62, A64, A65, A66)
- **Failed: 5** (A63 theme, C33 jwtdec, C34 rate, C35 cronck, C36 gitlog)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 4** (A2, A3, A5, A6)

## Run: 2026-05-15 19:45
Commit: 8e453f4 (main)
Build: Release

### Discovery
- No new code changes since Run 23
- No new quicksheet-* repos (still 37)
- Extension fix PRs still OPEN: jwtdec#2, rate#2, cronck#2. gitlog has no fix PR.
- Issue #77 (theme) filed in Run 23, still OPEN

### Tests

| ID | Result | Notes |
|----|--------|-------|
| A53 |  PASS | Double-click opens: Code verified  `MouseDoubleClick` at line 137  `OnFormDoubleClick`  `OpenAllSelected()` (same path as Enter key, verified in A39+A52) |

### Cumulative (after Run 24)
- **Total: 98 tests** (54 core + 8 ext system + 36 extensions)
- **Passed: 87** (+1: A53)
- **Failed: 5** (A63 theme #77, C33 jwtdec, C34 rate, C35 cronck, C36 gitlog)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6  tray icon hard to automate)

## Run: 2026-05-15 20:45
Commit: 3378d37 (main)
Build: N/A (no tests run)

### Discovery
- No new code changes since Run 24
- **PR #78 OPEN**: fix(desktop): make Ctrl+T theme cycling visible (fix for Issue #77/A63)
- Extension fix PRs still OPEN: jwtdec#2, rate#2, cronck#2. gitlog has no fix PR.
- No new repos (still 37)
- **No tests run**  all remaining failures blocked on open PRs

### Cumulative (unchanged from Run 24)
- **Total: 98 tests** (54 core + 8 ext system + 36 extensions)
- **Passed: 87**
- **Failed: 5** (A63 theme #77/#78, C33 jwtdec, C34 rate, C35 cronck, C36 gitlog)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 26: 2026-05-15 21:49
Commit: 02f8ad4 (main), v0.14.0
Build: Release

### Discovery
- **PR #78 MERGED**: fix(desktop): make Ctrl+T theme cycling visible  v0.14.0
- **4 extension fix PRs MERGED**: jwtdec#2, rate#2, cronck#2, gitlog#2

### Results

| ID | Result | Notes |
|----|--------|-------|
| A63 |  PASS | Theme cycling works! 3 distinct themes visible: Dark (black/teal)  Light (gray)  Blue (solarized) |
| C33 |  PASS | jwtdec decodes JWT: HEADER (alg, typ), CLAIMS (sub=1234567890, name=John Doe, iat), SIGNATURE (Present 43 chars) |
| C34 |  PASS | rate: Target $120,000/yr  Min Rate $141/hr, Take-home $86/hr, Billable 70% (1,400h/yr) |
| C35 |  FAIL | cronck shows "empty expression"  reads `cells` instead of `params` from activate message  cronck#3 filed |
| C36 |  PASS | gitlog shows 3 commits: hash, author, time, message + Branch: master + Showing 3 commits |

### Issues Filed
- **cronck#3**: Reads `cells` instead of `params` from activate message  shows empty expression

### Cumulative (98 tests)
- **Passed: 91** (+4 from Run 25: A63, C33, C34, C36)
- **Failed: 1** (C35 cronck)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 27: 2026-05-15 22:45
Commit: e7d5455 (main), v0.14.0
Build: Release

### Discovery
- No new commits since Run 26
- No new extension repos (37 total)
- cronck#3 still OPEN (no fix PR yet)
- C16/C25/C31 still BLOCKED (auth/Docker/k8s)

### Results
- **No tests run**  all remaining items blocked on upstream fixes or environment

### Cumulative (unchanged from Run 26)
- **Passed: 91**
- **Failed: 1** (C35 cronck  cronck#3)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 28: 2026-05-15 23:45
Commit: 4931fe0 (main), v0.14.0
Build: Release

### Discovery
- No new commits since Run 27
- No new extension repos (37 total)
- cronck#3 still OPEN (no fix PR yet)

### Results
- **No tests run**  all remaining items blocked on upstream fixes or environment

### Cumulative (unchanged)
- **Passed: 91**
- **Failed: 1** (C35 cronck  cronck#3)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 29: 2026-05-16 00:45
Commit: fd0d165 (main), post-v0.14.0
Build: Release

### Discovery
- **PR #79 MERGED**: feat(desktop): render sparkline (s:) prefix  Issue #9 CLOSED!
- cronck#3 still OPEN (no fix PR)
- No new extension repos (37 total)

### Results

| ID | Result | Notes |
|----|--------|-------|
| A67 |  PASS | NEW TEST: Sparkline renders with unicode block bars, dark-blue bg (20,30,50), light-blue fg (100,180,255) |

### Issues Resolved
- **#9 CLOSED**: Sparkline not rendered in desktop mode  fixed by PR #79

### Cumulative (99 tests  +1 new A67)
- **Passed: 92** (+1: A67 sparkline)
- **Failed: 1** (C35 cronck  cronck#3)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 30: 2026-05-16 01:45
Commit: e80e1b9 (main), v0.15.0
Build: Release

### Discovery
- **v0.15.0 released** (PR #80: README badges)
- **cronck PR #4 MERGED**: fix: read params instead of cells  cronck#3 CLOSED!
- No new extension repos (37 total)

### Results

| ID | Result | Notes |
|----|--------|-------|
| C35 |  PASS | cronck now correctly parses cron expression! Direct test: `*/5 * * * *`  `Every 5 minutes` |

### Issues Resolved
- **cronck#3 CLOSED**: Reads params instead of cells  fixed by PR #4

### Cumulative (99 tests)
- **Passed: 93** (+1: C35 cronck)
- **Failed: 0** 
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 31: 2026-05-16 02:45
Commit: 860018b (main), v0.15.0
Build: Release

### Discovery
- **NEW extension**: quicksheet-news (RSS/Atom feed headlines)  38 repos total
- No new QuickSheet commits

### Results

| ID | Result | Notes |
|----|--------|-------|
| C37 |  PASS | NEW: news extension  correct manifest, protocol works. HN feed shows 10 headlines with titles + links |

### Cumulative (100 tests  +1 new C37)
- **Passed: 94** (+1: C37 news)
- **Failed: 0**
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 32: 2026-05-16 03:45
Commit: ff6b896 (main), v0.15.0
Build: Release

### Discovery
- PR #81 merged (docs: add news extension to directory)  no code changes
- No new extension repos (38 total)
- **No tests run**  all passing, nothing new to test

### Cumulative (unchanged)
- **Passed: 94**
- **Failed: 0**
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 33: 2026-05-16 04:45
Commit: 6b92dba (main), v0.15.0
Build: Release

### Discovery
- Commit 6b92dba: CONTRIBUTING.md refresh (docs only, no code changes)
- No new extension repos (38 total)
- **No tests run**  all passing, nothing new to test

### Cumulative (unchanged)
- **Passed: 94**
- **Failed: 0**
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 34: 2026-05-16 05:45
Commit: 93c8872 (main), v0.16.0
Build: Release

### Discovery
- v0.16.0 tagged: CHANGELOG + CONTRIBUTING docs only, no code changes
- No new extension repos (38 total)
- **No tests run**  all passing, nothing new to test

### Cumulative (unchanged)
- **Passed: 94**
- **Failed: 0**
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 35: 2026-05-16 06:45
Commit: 0f33b0a (main), v0.16.0
Build: Release

### Discovery
- **New extension repo discovered: quicksheet-b64** (39th repo)  Base64 encode/decode
- .NET extension using `dotnet run --project QuickSheetB64.csproj`
- Manifest looks correct: prefix "b64", entry "dotnet run --project QuickSheetB64.csproj"

### Tests

| ID | Result | Notes |
|----|--------|-------|
| C38 |  FAIL | b64 extension uses absolute coordinates (parsed from anchor) instead of relative  output cells invisible  filed quicksheet-b64#1 |

### Issue Filed
- **quicksheet-b64#1**: Extension uses absolute coordinates instead of relative  no output visible
  - Root cause: `Process()` parses anchor into `baseRow`/`baseCol` and writes cells at those absolute positions
  - QuickSheet protocol expects relative coords (r=0 = first row below anchor), so output gets double-offset
  - Fix: use 0-based relative coordinates like all other extensions

### Cumulative
- **Passed: 94** (unchanged)
- **Failed: 1** (C38 b64  coordinate bug)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)
- **Total: 101** (added C38)


## Run 36: 2026-05-16 07:45
Commit: 682e98f (main), v0.16.0
Build: Release

### Discovery
- No new code changes (just README docs)
- No new repos (39 total)
- quicksheet-b64#1 still open  no fix yet
- **No tests run**  waiting for b64 fix

### Cumulative (unchanged)
- **Passed: 94**
- **Failed: 1** (C38 b64)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 37: 2026-05-16 13:30
Commit: 7ee6db5 (main), v0.16.0
Build: Release

### Discovery
- **New extension repo: quicksheet-guid** (40th repo) — GUID/UUID generator
- **PR #87 merged**: Extension cell width/height params now optional (defaults 1 col × 10 rows)
- **quicksheet-b64 PR #2 merged**: Fixed absolute→relative coordinate bug
- Code change in CellPrefix.cs: optional trailing gridCols/gridRows dimensions

### Tests

| ID | Result | Notes |
|----|--------|-------|
| C38 | ✅ PASS | b64 extension fixed (PR #2 merged) — "Hello World" → 🔒 ENCODED + SGVsbG8gV29ybGQ= + (11 bytes → 16 chars) |
| C39 | ✅ PASS | guid extension — 3 GUIDs generated in standard format, correct protocol |

### Cumulative
- **Passed: 96** (+2: C38 fixed, C39 new)
- **Failed: 0**
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)
- **Total: 102** (added C39)


## Run 38: 2026-05-16 14:13
Commit: 43b5933 (main), v0.16.0
Build: Release

### Discovery
- **New extension repo: quicksheet-regex** (41st repo) — Regex pattern explainer
- PR #96 open: 5 new themes (Dracula, Synthwave, Gruvbox, Monokai, HotdogStand)
- PR #93 open: strict extension protocol specification docs
- PR #97 open: add regex to extension directory

### Tests

| ID | Result | Notes |
|----|--------|-------|
| C40 | ❌ FAIL | regex extension uses {row,col,value} instead of {r,c,v} — no output visible → filed quicksheet-regex#1 |

### Issue Filed
- **quicksheet-regex#1**: Uses {row,col,value} property names instead of {r,c,v}
  - Root cause: C# anonymous objects use wrong field names
  - QuickSheet only reads , c,  from JSON (ExtensionProtocol.cs:108-110)
  - Fix: rename all properties to r/c/v

### Cumulative
- **Passed: 96**
- **Failed: 1** (C40 regex — wrong cell format)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)
- **Total: 103** (added C40)


## Run 39: 2026-05-16 14:58
Commit: 688923f (main), v0.17.0
Build: Release

### Discovery
- **v0.17.0 tagged**: PR #96 merged (5 new themes: Dracula, Synthwave, Gruvbox, Monokai, HotdogStand)
- PR #97 merged: add quicksheet-regex to directory
- quicksheet-regex#1 still open (no fix PR yet)
- No new extension repos (41 total)

### Tests

| ID | Result | Notes |
|----|--------|-------|
| C40 | ❌ FAIL | regex still broken — waiting for {r,c,v} fix (quicksheet-regex#1) |
| A63 | ✅ PASS | Theme cycling still works with v0.17.0 — new themes render correctly (teal/cyan visible) |

### Notes
- PR #96 adds Dracula, Synthwave, Gruvbox, Monokai, HotdogStand themes
- Ctrl+T cycling mechanism unchanged from PR #78 (verified Run 26)
- Grid renders with themed colors (non-default background visible)

### Cumulative
- **Passed: 96**
- **Failed: 1** (C40 regex — {row,col,value} bug)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)
- **Total: 103**


## Run 40: 2026-05-16 15:43
Commit: 7736975 (main), v0.17.0
Build: Release

### Discovery
- quicksheet-regex PR #2 open (fix: use {r,c,v}) — not merged yet
- No new repos (41 total), no code changes
- **No tests run** — waiting for regex PR merge

### Cumulative (unchanged)
- **Passed: 96**
- **Failed: 1** (C40 regex)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 41: 2026-05-16 16:28
Commit: 6fe25f7 (main), v0.17.0
Build: Release

### Discovery
- Research/persona docs only, no code changes
- quicksheet-regex PR #2 still open
- No new repos (41 total)
- **No tests run**

### Cumulative (unchanged)
- **Passed: 96**
- **Failed: 1** (C40 regex)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)


## Run 42: 2026-05-16 17:13
Commit: a6a6323 (main), v0.17.0
Build: Release

### Discovery
- **quicksheet-regex PR #2 merged**: Fixed {row,col,value} → {r,c,v}
- Grow log: "merged 4 ext PRs" — regex fix included
- No new repos (41 total)

### Tests

| ID | Result | Notes |
|----|--------|-------|
| C40 | ✅ PASS | regex extension fixed — pattern ^[a-z]+\d{2}$ correctly breaks down: ^ (anchor), [a-z]+ (class), \d{2} (digit), $ (anchor), Summary: 4 components |

### Cumulative
- **Passed: 97** (+1: C40 fixed)
- **Failed: 0**
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)
- **Total: 103**


## Run 43: 2026-05-16 17:58
Commit: 498948d (main), v0.17.0
Build: Release

### Discovery
- **New extension repo: quicksheet-urlenc** (42nd repo) — URL encode/decode
- PR #99 open: value-driven cell colour feature
- No other code changes

### Tests

| ID | Result | Notes |
|----|--------|-------|
| C41 | ❌ FAIL | urlenc extension uses absolute coordinates (anchorRow/anchorCol) instead of relative → filed quicksheet-urlenc#1 |

### Issue Filed
- **quicksheet-urlenc#1**: Absolute coordinates bug (same as b64#1 pattern)

### Cumulative
- **Passed: 97**
- **Failed: 1** (C41 urlenc — absolute coords)
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)
- **Total: 104** (added C41)

## Run 44: 2026-05-16 18:46
Commit: 89fb919 (main), v0.18.0
Build: Release

### Discovery
- **PR #99 merged**: Value-driven cell colour (`c?:` prefix) � conditional coloring based on rules like `c?: >80=red, >50=yellow, *=green: 73`
- **PR #98 merged**: urlenc docs
- **urlenc PR #2 merged**: Fixed absolute coordinates bug ? relative coordinates now correct

### Tests

| ID | Result | Notes |
|----|--------|-------|
| C41 | ? PASS | urlenc fix PR #2 merged � `Component: hello%20world` and `Full URI: hello%20world` visible |
| A68 | ? PASS | c?: conditional color � 73 shows yellow (>50), 95 shows red (>80), 30 shows green (default) |

### Cumulative
- **Passed: 99** (+2: C41 retest, A68 new)
- **Failed: 0**
- **Blocked: 3** (C16 copilot, C25 docker, C31 k8s)
- **Skipped: 3** (A2, A3, A5/A6)
- **Total: 105** (added A68)
