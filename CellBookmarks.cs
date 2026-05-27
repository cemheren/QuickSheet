namespace ExcelConsole;

/// <summary>
/// Manages up to 5 named cell bookmarks (slots 1-5).
/// Power users can mark frequently-visited cells and jump back instantly.
/// Bookmarks are persisted via a special config row in the CSV (like theme).
/// </summary>
public class CellBookmarks
{
    private const int SlotCount = 5;
    private const string ConfigPrefix = "cfg:bookmarks:";

    // Each slot stores (row, col) or null if unset.
    private readonly (int row, int col)?[] _slots = new (int, int)?[SlotCount];

    /// <summary>
    /// Set a bookmark at the given slot (0-based index, 0..4 for slots 1..5).
    /// </summary>
    public void Set(int slot, int row, int col)
    {
        if (slot < 0 || slot >= SlotCount) return;
        _slots[slot] = (row, col);
    }

    /// <summary>
    /// Get the bookmark at the given slot, or null if unset.
    /// </summary>
    public (int row, int col)? Get(int slot)
    {
        if (slot < 0 || slot >= SlotCount) return null;
        return _slots[slot];
    }

    /// <summary>
    /// Load bookmarks from a GridManager by scanning for the config cell.
    /// </summary>
    public void LoadFrom(GridManager grid)
    {
        for (int r = 0; r < grid.RowCount; r++)
            for (int c = 0; c < grid.ColumnCount; c++)
            {
                string val = grid.GetCellValue(r, c);
                if (val.StartsWith(ConfigPrefix, StringComparison.Ordinal))
                {
                    Parse(val[ConfigPrefix.Length..]);
                    return;
                }
            }
    }

    /// <summary>
    /// Persist bookmarks into the grid via a config cell.
    /// Uses the same pattern as ConfigCell (theme persistence).
    /// Writes to a known location (last row, last col) if not already present.
    /// </summary>
    public void PersistTo(GridManager grid)
    {
        string encoded = ConfigPrefix + Encode();

        // Find existing config cell
        for (int r = 0; r < grid.RowCount; r++)
            for (int c = 0; c < grid.ColumnCount; c++)
            {
                if (grid.GetCellValue(r, c).StartsWith(ConfigPrefix, StringComparison.Ordinal))
                {
                    grid.SetCellValue(r, c, encoded);
                    return;
                }
            }

        // Place in last row, second-to-last col (avoid collision with theme config)
        int targetRow = grid.RowCount - 1;
        int targetCol = Math.Max(0, grid.ColumnCount - 2);
        grid.SetCellValue(targetRow, targetCol, encoded);
    }

    private string Encode()
    {
        // Format: "r,c;r,c;r,c;r,c;r,c" — empty slots are "-"
        var parts = new string[SlotCount];
        for (int i = 0; i < SlotCount; i++)
            parts[i] = _slots[i] is var (r, c) ? $"{r},{c}" : "-";
        return string.Join(";", parts);
    }

    private void Parse(string data)
    {
        string[] parts = data.Split(';');
        for (int i = 0; i < Math.Min(parts.Length, SlotCount); i++)
        {
            if (parts[i] == "-") { _slots[i] = null; continue; }
            string[] rc = parts[i].Split(',');
            if (rc.Length == 2 && int.TryParse(rc[0], out int r) && int.TryParse(rc[1], out int c))
                _slots[i] = (r, c);
            else
                _slots[i] = null;
        }
    }
}
