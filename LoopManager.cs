namespace ExcelConsole;

/// <summary>
/// Manages L: (loop) cells that periodically activate a target cell.
/// Scans the grid for L: prefixed cells, parses their targets, and fires
/// an activation callback on the configured interval.
/// </summary>
internal sealed class LoopManager : IDisposable
{
    private GridManager _grid;
    private readonly Action<int, int> _activateCell;
    private readonly Dictionary<(int row, int col), LoopEntry> _entries = new();
    private readonly System.Threading.Timer _scanTimer;

    /// <summary>
    /// Creates a LoopManager that scans for L: cells and activates targets.
    /// </summary>
    /// <param name="grid">The grid to scan for L: cells.</param>
    /// <param name="activateCell">Callback to activate a cell (row, col), equivalent to Enter/double-click.</param>
    public LoopManager(GridManager grid, Action<int, int> activateCell)
    {
        _grid = grid;
        _activateCell = activateCell;
        // Scan every 30 seconds for new/changed L: cells
        _scanTimer = new System.Threading.Timer(_ => ScanGrid(), null,
            TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(30));
    }

    /// <summary>
    /// Scans the grid for L: cells and sets up or tears down timers as needed.
    /// </summary>
    public void ScanGrid()
    {
        var found = new HashSet<(int, int)>();

        for (int r = 0; r < _grid.RowCount; r++)
        {
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                string val = _grid.GetCellValue(r, c);
                if (!CellPrefix.IsLoop(val)) continue;

                var parsed = CellPrefix.ParseLoop(val);
                if (parsed == null) continue;

                var key = (r, c);
                found.Add(key);

                // Check if already tracked with same config
                if (_entries.TryGetValue(key, out var existing))
                {
                    if (existing.TargetRow == parsed.Value.row &&
                        existing.TargetCol == parsed.Value.col &&
                        existing.Minutes == parsed.Value.minutes)
                        continue; // unchanged

                    // Config changed, dispose old timer and recreate
                    existing.Timer.Dispose();
                    _entries.Remove(key);
                }

                // Create new entry
                var interval = TimeSpan.FromMinutes(parsed.Value.minutes);
                var entry = new LoopEntry
                {
                    TargetRow = parsed.Value.row,
                    TargetCol = parsed.Value.col,
                    Minutes = parsed.Value.minutes,
                    Timer = new System.Threading.Timer(_ =>
                    {
                        try { _activateCell(parsed.Value.row, parsed.Value.col); }
                        catch { }
                    }, null, interval, interval)
                };
                _entries[key] = entry;
            }
        }

        // Remove entries for cells that no longer have L: prefix
        var toRemove = _entries.Keys.Where(k => !found.Contains(k)).ToList();
        foreach (var key in toRemove)
        {
            _entries[key].Timer.Dispose();
            _entries.Remove(key);
        }
    }

    /// <summary>
    /// Updates the grid reference (e.g. after a resize/rebuild).
    /// Re-scans immediately to adjust timers to the new grid.
    /// </summary>
    public void UpdateGrid(GridManager grid)
    {
        _grid = grid;
        // Stop all existing timers and rescan with the new grid
        foreach (var entry in _entries.Values)
            entry.Timer.Dispose();
        _entries.Clear();
        ScanGrid();
    }

    public void Dispose()
    {
        _scanTimer.Dispose();
        foreach (var entry in _entries.Values)
            entry.Timer.Dispose();
        _entries.Clear();
    }

    private sealed class LoopEntry
    {
        public int TargetRow;
        public int TargetCol;
        public int Minutes;
        public System.Threading.Timer Timer = null!;
    }
}
