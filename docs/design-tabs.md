# Design: Virtual Tabs (Weekly Aging)

> Addresses [#158](https://github.com/cemheren/QuickSheet/issues/158).

## Summary

QuickSheet gets **virtual tabs** — multiple named sheets stored in one CSV file.
Cells age out weekly from the current tab into older tabs, keeping the active
desktop fresh while preserving history for search.

## Data model

```
TabManager
 ├── tabs: List<Tab>       (ordered, newest-last)
 ├── activeIndex: int      (index of the "Current" tab)
 └── config
      ├── maxTabs: int     (default 4)
      └── agingDay: DayOfWeek (default Monday)

Tab
 ├── name: string          ("Current", "Week -1", "Week -2", "Archive")
 └── grid: GridManager
```

Each `Tab` owns a `GridManager` instance with the same row/column dimensions.

## CSV persistence

Tabs are stored in a single CSV separated by a marker row:

```csv
cell,data,here
more,cells,
---TAB:Week -1---
old,data,from last week
---TAB:Week -2---
even,older,data
---TAB:Archive---
ancient,data,
```

- The first block (before any `---TAB:...---` row) is always the **Current** tab.
- Marker rows start with `---TAB:` and end with `---`. The text between is the tab name.
- `SaveToCsv` and `LoadFromCsv` in `GridManager` must be extended to recognize
  markers and delegate to/from `TabManager`.
- Backward-compatible: files without markers load as a single-tab sheet.

## Aging logic

On startup (or a configurable trigger), `TabManager.RunAging()`:

1. Check if a new week has started since last aging (compare stored
   `lastAgedDate` against today).
2. If yes:
   - Shift each tab's content one slot older: Archive ← Week -2 ← Week -1 ← Current.
   - **Only move "plain" cells** — cells with prefixes (`r:`, `ext:`, `i:`, `s:`,
     URLs, `{ref}`) stay pinned on the Current tab.
   - Clear the moved cells on Current (they now live in Week -1).
3. Store `lastAgedDate = today` in the CSV (as a metadata comment line at top,
   e.g., `#META:lastAged=2026-06-02`).

### What stays pinned (never ages)

| Prefix/Pattern       | Reason                         |
|----------------------|--------------------------------|
| `r: ...`            | Runnable command — user intent |
| `ext: ...`          | Extension binding              |
| `i: ...`            | Inline process                 |
| `s: ...`            | Sparkline definition           |
| `http://`, `https://` | Hyperlink                    |
| `{A1::C3}`          | Cell-range reference           |

## UX

### Tab indicator (bottom bar)

```
[ Archive | W-2 | W-1 | ●Current ]
```

- Shown in the status/bottom bar area.
- Active tab highlighted (bold or inverse).
- Fits in ≤40 chars for typical 4-tab setup.

### Switching tabs

| Key            | Action                |
|----------------|----------------------|
| `Ctrl+PgUp`   | Previous tab          |
| `Ctrl+PgDn`   | Next tab              |
| `Ctrl+T`      | Create new tab cycle  |

- When viewing a non-current tab, cells are **read-only** (or show a visual
  dimming) to prevent accidental edits to archived data.
- Search (`/`) spans all tabs; results indicate which tab the match is in.

### Manual tab creation

`Ctrl+T` triggers an immediate aging cycle regardless of the weekly timer.
This is the "fresh desktop" action from the issue.

## Implementation phases

1. **Phase 1 — Data layer** (~100 LOC)
   - `TabManager.cs`: holds `List<Tab>`, serializes/deserializes multi-tab CSV.
   - Extend `GridManager.SaveToCsv` / `LoadFromCsv` to call through `TabManager`.
   - Unit-testable without UI.

2. **Phase 2 — Tab switching UI** (~50 LOC)
   - Bottom-bar indicator rendering.
   - Key bindings for tab navigation.
   - Active tab swap in the host's render loop.

3. **Phase 3 — Aging logic** (~60 LOC)
   - `RunAging()` with pinned-cell detection.
   - `lastAgedDate` metadata persistence.
   - `Ctrl+T` manual trigger.

4. **Phase 4 — Cross-tab search** (~30 LOC)
   - Extend existing search to iterate all tabs.
   - Result display shows tab name.

## Open questions

- Should the user be able to **name** tabs freely, or is the week-based naming
  fixed? (Issue suggests week-based is sufficient.)
- Should the Archive tab have a max row limit, or grow unbounded?
- Should aging happen silently on startup, or prompt the user once?
- Should `Ctrl+T` be "force new week" or "add a named tab"?

## Compatibility

- Old CSV files (no markers) load normally as a single-tab sheet.
- If a user downgrades, the marker rows appear as visible cell content in row 1
  of a merged view — non-destructive but ugly. Document in release notes.
