---
name: test-desktop
description: >
  QA tester for QuickSheet desktop mode. Systematically tests all core features
  (navigation, editing, search, CSV, inline refs, commands) and all 16 extensions.
  Files GitHub issues with screenshots and debug logs for any bugs found.
  Maintains a persistent test-log.md for incremental testing across sessions.
  Use when user says "test desktop", "run tests", "QA", or invokes /test-desktop.
  Currently tracking 27 extensions (C1-C27) plus 48 core tests and 8 extension system tests.
---

# QuickSheet Desktop Mode Testing Skill

You are a QA tester for QuickSheet, a .NET 9 spreadsheet that replaces the Windows desktop wallpaper. Your job is to systematically test all desktop mode features and all extensions, filing GitHub issues for any bugs found.

## Discovery Phase (Run Periodically)

Before each test session, re-discover the codebase and extensions to catch new features/extensions that need testing:

### 1. Check for new code changes
```bash
cd <QuickSheet repo>
git pull
git --no-pager log --oneline -20  # recent commits
```

### 2. Scan for new features
```bash
# New cell prefixes
grep -rn "IsExtension\|IsInline\|IsCommand\|IsHyperlink\|StartsWith" CellPrefix.cs
# New keyboard shortcuts in DesktopForm
grep -n "case Keys\." Platform/Windows/DesktopForm.cs | head -50
# New files
git --no-pager diff --stat HEAD~10
```

### 3. Discover new extensions
```bash
gh repo list cemheren --json name,description --limit 50 | \
  jq '.[] | select(.name | startswith("quicksheet-"))'
```

Compare against test-log.md — any new `quicksheet-*` repos should be added to Group C.

### 4. Check extension docs for protocol changes
```bash
cat docs/extensions.md  # updated extension table
```

### 5. Update the skill
If new features or extensions are found:
1. Add new test cases to the appropriate group in `instructions.md`
2. Add new test IDs to `test-log.md` (with no result yet)
3. Commit and push the updated skill

### Discovery cadence
- Run discovery at the **start** of every test session
- If >5 commits since last run, do a full re-scan
- If new `quicksheet-*` repos appear, add them immediately

## Tools Available

This skill uses the **windows-mcp** MCP server for GUI automation:

- `windows-mcp-Snapshot` — capture desktop state + screenshot (`useVision: true`)
- `windows-mcp-Click` — click at coordinates `[x, y]`
- `windows-mcp-Type` — type text at coordinates (use `clear: true` to replace, `pressEnter: true` to submit)
- `windows-mcp-Move` — move mouse / drag
- `windows-mcp-MultiEdit` — type into multiple fields
- `windows-mcp-MultiSelect` — Ctrl+click multiple items
- `windows-mcp-Clipboard` — get/set clipboard
- `windows-mcp-App` — launch/switch/resize windows

Also use PowerShell for:
- Building and launching QuickSheet
- Reading debug logs and CSV files
- Taking screenshots via `System.Drawing` (for issue attachments)
- Filing issues via `gh` CLI

### Workflow for each test

1. Use `windows-mcp-Snapshot` with `useVision: true` to see current state
2. Use `windows-mcp-Click` / `windows-mcp-Type` to interact
3. Use `windows-mcp-Snapshot` again to verify result
4. Log pass/fail in test-log.md

## Prerequisites

- .NET 9 SDK installed
- Git available on PATH
- `gh` CLI authenticated (for filing issues)
- GitHub Copilot CLI authenticated (only needed for the `copilot` extension test)

## How to Run

1. **Build** QuickSheet first to catch compile errors
2. **Launch** in desktop mode: `dotnet run -c Release --project ExcelConsole.csproj -- --desktop`
3. Walk through each test group (A → B → C) in order
4. For each test: perform the steps, verify the expected outcome, log pass/fail
5. For any failure: file a GitHub issue using the templates below

## Important Notes

- QuickSheet runs as a borderless WinForms window covering the desktop working area (behind the taskbar)
- It hides from Alt+Tab (WS_EX_TOOLWINDOW style)
- **Ctrl+Q** quits the app — use this to exit between test sessions
- Autosave goes to `Desktop/autosave.csv` (not AppData in desktop mode)
- Extensions install to `%APPDATA%/QuickSheet/extensions/`
- Extension debug log: `~/.quicksheet/extensions/debug.log`
- To start fresh: delete `Desktop/autosave.csv` and the extensions directory before testing

## Test Group A: Core Desktop Mode (46 tests)

### Launch & Window Management

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A1 | Launch & render | `dotnet run -c Release --project ExcelConsole.csproj -- --desktop` | Black grid covering working area, tray icon visible, console window hidden |
| A2 | Win+D show desktop | Press Win+D while other windows open | QuickSheet becomes topmost (visible above desktop icons) |
| A3 | Alt+Tab hidden | Press Alt+Tab | QuickSheet NOT in the task switcher list |
| A4 | Click to focus | Click on the grid | Grid receives keyboard focus |
| A5 | Tray icon menu | Right-click system tray icon | Context menu shows "Save" and "Exit" |
| A6 | Tray Exit | Click "Exit" in tray menu | App exits cleanly |
| A7 | Quit (Ctrl+Q) | Press Ctrl+Q | App exits cleanly, data autosaved |

### Navigation

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A8 | Arrow key nav | Press Up/Down/Left/Right | Cursor moves cell by cell, highlight follows |
| A9 | Tab navigation | Press Tab | Moves right; wraps to next row at end |
| A10 | Mouse click select | Click on a cell | Cursor jumps to clicked cell |
| A11 | Bounds clamping | Navigate to top-left corner, press Up and Left | Stays at (0,0), no crash |

### Cell Editing

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A12 | Direct typing | Navigate to empty cell, type `hello world` | Characters appear in cell one by one |
| A13 | F2 edit mode | Press F2 on a cell with content | Status bar shows "Edit: ..." with cursor, Left/Right move cursor |
| A14 | F2 commit (Enter) | In edit mode, modify text, press Enter | Cell value updated |
| A15 | F2 cancel (Escape) | In edit mode, modify text, press Escape | Cell value unchanged |
| A16 | Backspace | Type text, press Backspace | Last character removed |
| A17 | Delete key | Select cell with content, press Delete | Cell content cleared |
| A18 | Paste in edit mode | F2, then Ctrl+V with clipboard text | Text inserted at cursor position in edit bar |

### Selection

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A19 | Shift+Arrow extend | Shift+Down×3, Shift+Right×2 | Rectangle selection highlighted (blue tint) |
| A20 | Ctrl+Click toggle | Click cell, Ctrl+Click another, Ctrl+Click first again | Second added, first removed from selection |
| A21 | Mouse drag | Click and drag across cells | Rectangle selection formed in real-time |
| A22 | Escape clears | Make multi-selection, press Escape | All selection cleared |
| A23 | Delete multi-select | Select multiple cells, press Delete | All selected cells cleared |

### Copy / Cut / Paste

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A24 | Ctrl+C single | Type "test", Ctrl+C | Cell value on system clipboard |
| A25 | Ctrl+C multi-select | Select 3 cells, Ctrl+C | All values on clipboard (newline-separated) |
| A26 | Ctrl+X cut | Type "test", Ctrl+X | Cell cleared, value on clipboard |
| A27 | Ctrl+V paste | Copy text, navigate elsewhere, Ctrl+V | Text pasted |
| A28 | Multi-line paste | Copy "line1\nline2\nline3" externally, Ctrl+V | Lines fill consecutive rows in same column |
| A29 | Ctrl+Shift+C resolved | Create `i: A1` ref, Ctrl+Shift+C on it | Resolved value (not "i: A1") copied |

### Row Operations

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A30 | Ctrl+D delete row | Put data in row 3, Ctrl+D on row 3 | Row deleted, rows below shift up |
| A31 | Ctrl+O insert below | Ctrl+O on row 3 | Blank row inserted at row 3, data shifts down |
| A32 | Ctrl+P shift up | Ctrl+P on row 3 | Rows shift up (row 3 gets row 4's content) |

### Search

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A33 | Ctrl+F enter search | Press Ctrl+F | Status bar shows "Find: │" prompt |
| A34 | Type search term | Type "hello", press Enter | Matching cells highlighted yellow, cursor jumps to first match |
| A35 | Enter = next match | Press Enter again | Moves to next match (wraps around) |
| A36 | Shift+Enter = prev | Press Shift+Enter | Moves to previous match |
| A37 | Escape exits search | Press Escape | Search highlights cleared, back to normal mode |
| A38 | No matches | Search for nonexistent term | Status bar shows "no matches" |

### Cell Prefixes & Features

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A39 | Hyperlink (http) | Type `https://github.com`, press Enter | Default browser opens URL; cell has purple background |
| A40 | Command (r:) | Type `r: notepad`, press Enter | Notepad launches; cell has yellow/amber background |
| A41 | Inline ref (i:) | A1=`r: echo hello`, B1=`i: A1` | B1 shows "hello" with teal background |
| A42 | Inline span | A1=`r: echo "multi line"`, B1=`i: A1,3,5` | Multi-cell overlay with border, content, and cell ref label |
| A43 | Inline rerun | Press Enter on an i: cell pointing to a command | Process restarts (output refreshes) |
| A44 | Cell range {A1::B2} | Put data in A1-B2, type `{A1::B2}` in C1 | C1 shows expanded cell values (tab/newline separated) |
| A45 | F1 resolve toggle | Select cell with i: ref, press F1 | Status bar toggles between raw "i: A1" and resolved value |

### Calculations

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A46 | Column sum (Σ) | Put 10, 20, 30 in column A, select any cell in A | Status bar shows "ΣA = 60" |
| A47 | Row product (Π) | Put 2, 3, 4 in row 1, select any cell in row 1 | Status bar shows "Π1 = 24" |

### Grid & Display

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A48 | F3 rebuild | Press F3 | Grid recalculated, data preserved, desktop files re-populated |
| A49 | F4 shrink columns | Press F4 | Column width decreases (min 8), more columns visible |
| A50 | F5 expand columns | Press F5 | Column width increases (max 40), fewer columns visible |
| A51 | Desktop files shown | Check rightmost column(s) | Desktop files/folders listed, folders have 📁 prefix |
| A52 | Open desktop file | Select a desktop file, press Enter | File opens with default app |
| A53 | Double-click opens | Double-click hyperlink/command/file | Target opens |

### Persistence

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A54 | Autosave | Make changes, wait >5s | `Desktop/autosave.csv` updated |
| A55 | Ctrl+S save | Press Ctrl+S | Tray balloon "Saved to ..." appears |
| A56 | Load CSV | Launch with `-- --desktop test.csv` | Data from test.csv visible in grid |
| A57 | CSV merge | Edit CSV externally while app running, wait ~60s | External changes merged in |
| A58 | Conflict marker | Edit same cell locally and externally, wait for merge | Cell shows `c: external(local)` with red background |

### Find & Replace

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A59 | Ctrl+R find/replace | Press Ctrl+R, enter find term, enter replace term | Matching cells found and replaced (like console mode) |
| A60 | Ctrl+B sort column | Select column, press Ctrl+B | Column sorted alphabetically |

### Cell Color Prefix

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A61 | c:color: prefix | Type `c:red: URGENT` in a cell | Cell background turns red, text shows "URGENT" only |

### Missing Desktop Shortcuts

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| A62 | Ctrl+G go to cell | Press Ctrl+G | Status bar shows "Go to:" prompt (like console mode) |
| A63 | Ctrl+T cycle theme | Press Ctrl+T | Theme cycles (Dark → Light → Solarized → Nord → …) |
| A64 | Ctrl+H help overlay | Press Ctrl+H | Help overlay shows/hides |
| A65 | Ctrl+Z undo | Press Ctrl+Z | Last action undone |
| A66 | Ctrl+Y redo | Press Ctrl+Y | Last undone action redone |

---

## Test Group B: Extension System (8 tests)

| ID | Test | Steps | Expected |
|----|------|-------|----------|
| B1 | Install from GitHub | Type `ext: github:cemheren/quicksheet-weather` | Extension cloned, process launched, cell shows green (running) |
| B2 | Bad repo install | Type `ext: github:cemheren/nonexistent-repo` | Cell appended with `[install failed]` |
| B3 | No manifest repo | Type `ext: github:cemheren/1brc` | Cell appended with `[bad manifest]` |
| B4 | Prefix activation | After B1, type `wthr: Seattle,2,7` | Weather data fills cells below the prefix cell |
| B5 | Reactivate (Enter) | Press Enter on a prefix cell | Data refreshes |
| B6 | Cell clear deactivates | Clear a prefix cell | Extension receives deactivate; output cells may clear |
| B7 | Extension color coding | Observe ext: and prefix: cell backgrounds | Loading=distinct, Running=green-ish, Error=red-ish |
| B8 | F3 rebuild with extensions | Press F3 after installing extensions | Extensions re-scanned, prefix cells re-activated |

---

## Test Group C: Individual Extensions (19 extensions)

First install each extension via its `ext:` cell. Then test the prefix.

| ID | Repo | Install cell | Test cell | Expected |
|----|------|-------------|-----------|----------|
| C1 | quicksheet-weather | `ext: github:cemheren/quicksheet-weather` | `wthr: Seattle,2,7` | 7 rows: day abbreviations + weather emoji + temps |
| C2 | quicksheet-todo | `ext: github:cemheren/quicksheet-todo` | `todo: list,1,5` then `todo: add Buy milk,1,3` then `todo: list,1,5` | Task appears, list shows it |
| C3 | quicksheet-sysmon | `ext: github:cemheren/quicksheet-sysmon` | `sys: all,2,4` | 4 rows: CPU/RAM/disk/uptime with color bars, updates every 2s |
| C4 | quicksheet-pomodoro | `ext: github:cemheren/quicksheet-pomodoro` | `pomo: 1,1,2` | 1-min countdown with progress bar, updates every 1s |
| C5 | quicksheet-stock-ext | `ext: github:cemheren/quicksheet-stock-ext` | `stock: AAPL,1,3` | Price + change + date |
| C6 | quicksheet-price-ext | `ext: github:cemheren/quicksheet-price-ext` | `price: btc,1,2` | Bitcoin price + 24h change % |
| C7 | quicksheet-define-ext | `ext: github:cemheren/quicksheet-define-ext` | `def: laconic,1,4` | Word + part-of-speech definitions |
| C8 | quicksheet-thes-ext | `ext: github:cemheren/quicksheet-thes-ext` | `thes: happy,1,6` | Word + synonyms |
| C9 | quicksheet-cite-ext | `ext: github:cemheren/quicksheet-cite-ext` | `cite: 10.1145/3623476.3623525,1,4` | Citation: authors, title, venue, DOI |
| C10 | quicksheet-ping-ext | `ext: github:cemheren/quicksheet-ping-ext` | `ping: https://github.com,1,3` | ✓ URL + 200 OK + latency ms |
| C11 | quicksheet-mxck-ext | `ext: github:cemheren/quicksheet-mxck-ext` | `mxck: github.com,1,5` | MX records sorted by priority |
| C12 | quicksheet-mortgage-ext | `ext: github:cemheren/quicksheet-mortgage-ext` | `mort: 500000,6.5,30,1,4` | Loan summary + monthly + interest + total |
| C13 | quicksheet-tls-ext | `ext: github:cemheren/quicksheet-tls-ext` | `tls: github.com,1,4` | Host:443 + expiry days + issuer + CN |
| C14 | quicksheet-grav-ext | `ext: github:cemheren/quicksheet-grav-ext` | `grav: test@example.com,1,4` | Email + name + location + avatar URL |
| C15 | quicksheet-1099-ext | `ext: github:cemheren/quicksheet-1099-ext` | `1099: 80000,1,5` | Income + SE tax + quarterly + federal note + disclaimer |
| C16 | quicksheet-copilot-ext | `ext: github:cemheren/quicksheet-copilot-ext` | `copilot: list 3 colors,2,3` | AI response in grid cells (requires auth) |
| C17 | quicksheet-cal | `ext: github:cemheren/quicksheet-cal` | `cal: week,1,8` | Upcoming calendar events for next 7 days |
| C18 | quicksheet-fx | `ext: github:cemheren/quicksheet-fx` | `fx: 1000,USD,EUR,1,3` | Currency conversion: amount + rate + converted value (ECB rates) |
| C19 | quicksheet-qtr | `ext: github:cemheren/quicksheet-qtr` | `qtr: 2026,1,5` | Tax year + 4 quarterly deadlines with countdown |
| C20 | quicksheet-budget | `ext: github:cemheren/quicksheet-budget` | `budget: Groceries,500,350,1,3` | Category + progress bar + spent/budget + remaining |
| C21 | quicksheet-hntop | `ext: github:cemheren/quicksheet-hntop` | `hntop: 5,1,6` | Top 5 HN stories with scores and comment counts |
| C22 | quicksheet-apistatus | `ext: github:cemheren/quicksheet-apistatus` | `apistatus: github,npm,cloudflare,1,6` | Service status: emoji + name + operational/minor + description |
| C23 | quicksheet-ghpr | `ext: github:cemheren/quicksheet-ghpr` | `ghpr: cemheren/QuickSheet,1,5` | GitHub PRs needing attention (requires `gh` CLI auth) |
| C24 | quicksheet-portck | `ext: github:cemheren/quicksheet-portck` | `portck: 80,443,8080,1,5` | TCP port status: open/closed for each port |
| C25 | quicksheet-docker | `ext: github:cemheren/quicksheet-docker` | `docker: all,1,5` | Docker container status dashboard (requires Docker Desktop) |
| C26 | quicksheet-gitst | `ext: github:cemheren/quicksheet-gitst` | `gitst: .,1,5` | Git repo status: branch, clean/dirty, stash count, last commit |
| C27 | quicksheet-cntdn | `ext: github:cemheren/quicksheet-cntdn` | `cntdn: 2026-12-25,1,3` | Countdown to date: days, hours, minutes remaining with progress |
| C28 | quicksheet-worldtm | `ext: github:cemheren/quicksheet-worldtm` | `worldtm: London,Tokyo,NY,1,5` | Multi-timezone world clock with current times |
| C29 | quicksheet-mileage-ext | `ext: github:cemheren/quicksheet-mileage-ext` | `mileage: 1000,1,4` | IRS standard mileage rate calculation |
| C30 | quicksheet-margin-ext | `ext: github:cemheren/quicksheet-margin-ext` | `margin: 100000,60000,25000,1,5` | Break-even analysis: CM per unit, break-even units, revenue |
| C31 | quicksheet-k8s | `ext: github:cemheren/quicksheet-k8s` | `k8s: default,1,5` | Kubernetes pod status from kubeconfig (requires kubectl) |
| C32 | quicksheet-depr-ext | `ext: github:cemheren/quicksheet-depr-ext` | `depr: 50000 5 straight 5000,1,10` | Straight-line depreciation schedule: cost, salvage, yearly expense |
| C33 | quicksheet-jwtdec | `ext: github:cemheren/quicksheet-jwtdec` | `jwtdec: <jwt-token>,1,8` | Decoded JWT header + claims in grid cells |
| C34 | quicksheet-rate | `ext: github:cemheren/quicksheet-rate` | `rate: 120000,1,6` | Freelance rate calculator: min rate, take-home, billable hours |
| C35 | quicksheet-cronck | `ext: github:cemheren/quicksheet-cronck` | `cronck: */5 * * * *,1,3` | Human-readable cron schedule description |
| C36 | quicksheet-gitlog | `ext: github:cemheren/quicksheet-gitlog` | `gitlog: 5,1,6` | Recent 5 git commits from current repo |

---

## Issue Filing

When a test fails, file a GitHub issue. Follow these rules strictly:

### Rules

1. **Always include a screenshot** — capture the desktop (QuickSheet IS the desktop) showing the failure. Upload using `gh issue create` with the image attached or reference a screenshot URL.
2. **Include all debug details** — extension logs (`~/.quicksheet/extensions/debug.log`), stderr output, autosave CSV state, process status. The more context, the better.
3. **Scrub sensitive data** — remove any file paths containing usernames (replace with `~` or `%USERPROFILE%`), API keys, personal file names, or other PII before including in issues.
4. **Don't open duplicate issues** — before filing, check existing issues with `gh issue list --repo <repo> --state open` and search closed issues too. If a related issue exists, comment on it instead.

### Taking Screenshots

```powershell
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$bmp = New-Object System.Drawing.Bitmap($screen.Bounds.Width, $screen.Bounds.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Bounds.Location, [System.Drawing.Point]::Empty, $screen.Bounds.Size)
$g.Dispose()
$bmp.Save("screenshot.png", [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
```

### For core bugs (QuickSheet repo):
```bash
# First check for duplicates
gh issue list --repo cemheren/QuickSheet --state all --search "<keywords>"

# Then file with screenshot
gh issue create --repo cemheren/QuickSheet \
  --title "[Test] <Short description of failure>" \
  --body "## Test ID: <ID>

**Steps to reproduce:**
<exact steps>

**Expected:**
<what should happen>

**Actual:**
<what actually happened>

**Debug details:**
\`\`\`
<extension debug log, stderr output, relevant CSV state>
\`\`\`

**Screenshot:**
<attach screenshot showing the failure>

**Build:** Desktop mode, Release config, .NET 9
**OS:** Windows
**Commit:** $(git rev-parse --short HEAD)"
```

### For extension bugs (extension repo):
```bash
# First check for duplicates
gh issue list --repo cemheren/<extension-repo> --state all --search "<keywords>"

# Then file with screenshot
gh issue create --repo cemheren/<extension-repo> \
  --title "[Test] <Short description of failure>" \
  --body "## Test ID: <ID>

**Steps to reproduce:**
1. Install extension: \`ext: github:cemheren/<repo>\`
2. Activate: \`<prefix>: <params>\`
3. <what went wrong>

**Expected:**
<what should happen>

**Actual:**
<what actually happened>

**Debug details:**
\`\`\`
<~/.quicksheet/extensions/debug.log content>
<extension stderr output if available>
<relevant autosave.csv rows>
\`\`\`

**Screenshot:**
<attach screenshot showing the failure>

**QuickSheet version:** latest main (commit $(git rev-parse --short HEAD))
**Extension version:** latest main"
```

## Test Log

Maintain a persistent test log at `.agents/skills/test-desktop/test-log.md`. This enables incremental testing — you don't need to re-run tests that already passed.

### Log format

```markdown
# QuickSheet Test Log

## Run: YYYY-MM-DD HH:MM
Commit: <short hash>
Build: Release / Debug

| ID | Result | Notes |
|----|--------|-------|
| A1 | ✅ PASS | Window launched, tray icon visible |
| A2 | ❌ FAIL | Issue #42 filed |
| A3 | ⏭️ SKIP | Requires manual Alt+Tab verification |
| C16 | ⚠️ BLOCKED | Copilot CLI not authenticated |
```

### Rules for the log

1. **Append, don't overwrite** — each test run adds a new `## Run:` section
2. **Only re-test** items that previously failed, were skipped, or haven't been tested yet
3. **Link issues** — when you file an issue, add `→ issue #N` to the Notes column
4. **Mark regressions** — if a previously-passing test now fails, prefix with `🔄 REGRESSION`
5. **Record the commit hash** so results can be tied to a specific codebase state

### Reading the log

Before running tests, read `test-log.md` to determine what still needs testing:
- All `❌ FAIL` items: re-test to check if fixed
- All `⏭️ SKIP` items: attempt if conditions now allow
- Any test IDs not in the log: run them
- All `✅ PASS` items: skip unless the relevant code changed since that run

## Execution Procedure

1. **Read test-log.md** to see what's already been tested and what needs re-testing
2. **Clean state** (only if doing a full run): Delete `Desktop/autosave.csv` and `%APPDATA%/QuickSheet/extensions/`
3. **Build**: `dotnet build -c Release ExcelConsole.csproj`
4. **Launch**: `dotnet run -c Release --project ExcelConsole.csproj -- --desktop`
5. **Run pending tests** — Group A → B → C, skipping already-passed tests
6. **Update test-log.md** after each batch
7. **File issues** for failures (with screenshots + debug logs)
8. **Commit test-log.md** after each session
9. **Report** summary: total tests, passed, failed, skipped, issues filed
