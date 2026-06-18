# Export Formats

QuickSheet can export your spreadsheet data in multiple formats from the command line — no GUI needed.
All exports are headless (no TUI opens) and support stdout (`-`) for piping into other tools.

## CSV (native)

CSV is QuickSheet's native persistence format. Every autosave and manual save produces a standard CSV file.

```bash
# The grid autosaves to:
#   Windows: %APPDATA%/ExcelConsole/autosave.csv
#   Linux:   ~/.config/ExcelConsole/autosave.csv

# Desktop mode autosaves to its CSV file every 5 seconds.
```

## Markdown table (`--export-md`)

Convert a CSV file into a GitHub-flavored Markdown table.

```bash
# Export to a file
dotnet run --project ExcelConsole.csproj -- data.csv --export-md table.md

# Pipe to stdout (for clipboard, another tool, etc.)
dotnet run --project ExcelConsole.csproj -- data.csv --export-md -
```

**Example input** (`tasks.csv`):

```csv
Task,Status,Priority
Fix login bug,Done,High
Add dark mode,In Progress,Medium
Write docs,Pending,Low
```

**Output** (`tasks.md`):

```markdown
| Task | Status | Priority |
|---|---|---|
| Fix login bug | Done | High |
| Add dark mode | In Progress | Medium |
| Write docs | Pending | Low |
```

## HTML table (`--export-html`)

Convert a CSV file into a self-contained, styled HTML page. The output uses a dark theme
matching QuickSheet's aesthetic and includes:

- **Auto-linked URLs** — cells starting with `http://` or `https://` become clickable links
- **Right-aligned numbers** — numeric cells use tabular numerals and right alignment
- **Alternating row colors** and hover effects
- **QuickSheet backlink** in the footer

```bash
# Export to a file
dotnet run --project ExcelConsole.csproj -- data.csv --export-html report.html

# Pipe to stdout
dotnet run --project ExcelConsole.csproj -- data.csv --export-html -

# Pipe to clipboard (Windows)
dotnet run --project ExcelConsole.csproj -- data.csv --export-html - | clip

# Open directly in browser (Windows)
dotnet run --project ExcelConsole.csproj -- data.csv --export-html report.html && start report.html
```

## JSON array (`--export-json`)

Convert a CSV file into a JSON array of objects. The first row is used as keys;
numeric cells are emitted as numbers (not strings).

```bash
# Export to a file
dotnet run --project ExcelConsole.csproj -- data.csv --export-json data.json

# Pipe to stdout
dotnet run --project ExcelConsole.csproj -- data.csv --export-json -

# Pipe into jq for filtering / transformation
dotnet run --project ExcelConsole.csproj -- data.csv --export-json - | jq '.[] | select(.Priority == "High")'
```

**Example input** (`tasks.csv`):

```csv
Task,Status,Priority,Hours
Fix login bug,Done,High,3
Add dark mode,In Progress,Medium,8
Write docs,Pending,Low,2
```

**Output**:

```json
[
  {"Task":"Fix login bug","Status":"Done","Priority":"High","Hours":3},
  {"Task":"Add dark mode","Status":"In Progress","Priority":"Medium","Hours":8},
  {"Task":"Write docs","Status":"Pending","Priority":"Low","Hours":2}
]
```

## SQL (`--export-sql`)

Convert a CSV file into SQL `CREATE TABLE` + `INSERT` statements, compatible with SQLite, PostgreSQL, and MySQL.

```bash
# Export to a file
dotnet run --project ExcelConsole.csproj -- data.csv --export-sql schema.sql

# Pipe to stdout
dotnet run --project ExcelConsole.csproj -- data.csv --export-sql -

# Custom table name (default: "data")
dotnet run --project ExcelConsole.csproj -- data.csv --export-sql - --table tasks

# Load directly into SQLite
dotnet run --project ExcelConsole.csproj -- data.csv --export-sql - | sqlite3 mydb.sqlite
```

**Example input** (`tasks.csv`):

```csv
Task,Status,Priority,Hours
Fix login bug,Done,High,3
Add dark mode,In Progress,Medium,8
Write docs,Pending,Low,2
```

**Output** (`--table tasks`):

```sql
CREATE TABLE IF NOT EXISTS "tasks" ("Task" TEXT, "Status" TEXT, "Priority" TEXT, "Hours" TEXT);

INSERT INTO "tasks" ("Task", "Status", "Priority", "Hours") VALUES ('Fix login bug', 'Done', 'High', 3);
INSERT INTO "tasks" ("Task", "Status", "Priority", "Hours") VALUES ('Add dark mode', 'In Progress', 'Medium', 8);
INSERT INTO "tasks" ("Task", "Status", "Priority", "Hours") VALUES ('Write docs', 'Pending', 'Low', 2);
```

Numeric values are inserted without quotes. Empty cells become `NULL`. Column names with special characters are double-quoted per SQL standard.

## Piping and composition

All export modes support `-` as the output path, writing to stdout instead of a file.
This makes QuickSheet composable with other CLI tools:

```bash
# CSV → Markdown → copy to clipboard (macOS)
dotnet run --project ExcelConsole.csproj -- data.csv --export-md - | pbcopy

# CSV → JSON → pipe into jq
dotnet run --project ExcelConsole.csproj -- data.csv --export-json - | jq '.[0]'

# CSV → SQL → load into SQLite
dotnet run --project ExcelConsole.csproj -- data.csv --export-sql - | sqlite3 mydb.sqlite

# CSV → HTML → serve with Python
dotnet run --project ExcelConsole.csproj -- data.csv --export-html - > /tmp/report.html
python3 -m http.server -d /tmp 8080

# Chain with jq, awk, or other tools
dotnet run --project ExcelConsole.csproj -- data.csv --export-md - | head -5
```

## Format comparison

| Feature | CSV | Markdown | HTML | JSON | SQL |
|---|---|---|---|---|---|
| Human readable | ○ | ● | ● | ○ | ○ |
| Machine parseable | ● | ○ | ○ | ● | ● |
| Styled output | — | — | ● | — | — |
| Clickable URLs | — | ● (on GitHub) | ● | — | — |
| Numeric alignment | — | — | ● | ● (native) | ● (native) |
| Embeddable | — | ● (README/wiki) | ● (browser) | ● (APIs/scripts) | ● (databases) |
| Round-trip editable | ● | — | — | — | — |
| jq / tool-friendly | — | — | — | ● | — |
| DB-loadable | — | — | — | — | ● |

---

*See also: [Keyboard Shortcuts](keyboard-shortcuts.md) · [Dashboard Recipes](recipes.md) · [60-Second Tour](tour.md)*
