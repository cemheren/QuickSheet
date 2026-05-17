# QuickSheet for the csvkit / Miller / qsv crowd

If you already use [csvkit](https://csvkit.readthedocs.io/),
[Miller](https://miller.readthedocs.io/), [xsv](https://github.com/BurntSushi/xsv), or
[qsv](https://github.com/dathere/qsv) — QuickSheet is the *editing* surface on top of
the same plain CSV file. It does **not** try to replace your CLI tools; it complements
them. Pipe a transform out of `mlr` / `csvkit`, then edit / annotate the result in
QuickSheet.

```bash
# csvkit pipeline → QuickSheet (--export-md is the drop-in replacement for csvkit's csvlook → markdown)
csvjoin -c user_id users.csv orders.csv | csvgrep -c status -m active > active.csv
dotnet run --project ExcelConsole.csproj -- active.csv --export-md report.md
```

## The one feature CLI-CSV users care about: headless `--export-md`

QuickSheet's `--export-md` flag is the single most "drop-in replacement" feature in the
project: convert any CSV to a GitHub-flavoured Markdown table without launching the UI.
Useful for one-shot reports, CI pipelines, paste-into-Slack moments.

```bash
# Read mydata.csv, write a markdown table to mydata.md (no UI launched)
dotnet run --project ExcelConsole.csproj -- mydata.csv --export-md mydata.md

# Write to stdout instead — useful in pipelines
dotnet run --project ExcelConsole.csproj -- mydata.csv --export-md -
```

The flag uses a headless GridManager: no Console UI, no `--desktop` window, no
autosave. Trailing empty rows/columns are trimmed; `|` characters in cell content are
escaped; the first row becomes the header.

## What QuickSheet adds beyond the CLI tools

The CLI tools are great at *transforming* CSV. QuickSheet is good at *living with* it:

- Edit cells in place (Tad and Modern CSV are the closest peers — but neither runs as
  your wallpaper).
- Σ / Π in the status bar — quick sums and products without a formula language.
- Sparklines (`s: 1,2,3,...`) or unicode block bars from a range (`s: A1::A10`).
- Runnable cells (`r: code .`) for per-row commands.
- Stays open between sessions — autosave every 5 seconds.

## What QuickSheet doesn't try to do

- **No SQL.** If you want `WHERE` and `GROUP BY`, use `mlr --csv put '...'` or `qsv sql`.
  QuickSheet is for the *editing* moment, not the *query* moment.
- **No statistics.** csvkit's `csvstat` or `mlr stats1` does this better.
- **No type inference.** Cells are strings; you parse them with a CLI tool when needed.
- **No JSON / Parquet.** CSV in, CSV out. Use `mlr --c2j` to bridge.

## Recipe: open last week's deploys, annotate, export

```bash
# Pull deploys from your tool, filter to last week
deploy-cli list --json | mlr --ijson --ocsv cat \
  | mlr --csv filter '${started_at} >= "2026-05-10T00:00:00Z"' > deploys.csv

# Open in QuickSheet, add a column with notes per row
dotnet run --project ExcelConsole.csproj -- deploys.csv

# (You edit, you save, autosave handles the rest.)

# When you're done, export to Markdown and paste into the postmortem doc
dotnet run --project ExcelConsole.csproj -- deploys.csv --export-md postmortem-deploys.md
```

## See also

- [docs/tour.md](tour.md) — 60-second tour of cell prefixes (`r:`, `i:`, `s:`, …).
- [README.md](../README.md) — the wallpaper-mode pitch.
- [docs/extensions.md](extensions.md) — extension protocol + directory.
