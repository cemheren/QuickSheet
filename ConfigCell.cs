namespace ExcelConsole;

/// <summary>
/// Helpers for the <c>config:</c> cell prefix — persists UI state (currently the theme name)
/// inside the CSV itself rather than a sidecar file. Design decision per PR #130:
/// "I don't want it to be a file, it should be a value with a prefix."
/// </summary>
public static class ConfigCell
{
    /// <summary>
    /// Scan grid for the first <c>config:</c> cell and apply known settings (theme=…).
    /// Silently no-ops if no config cell is found or values are unknown.
    /// </summary>
    public static void LoadAndApply(GridManager grid)
    {
        var loc = Find(grid);
        if (loc == null) return;
        var map = CellPrefix.ParseConfig(grid.GetCellValue(loc.Value.row, loc.Value.col));
        if (map.TryGetValue("theme", out var themeName))
            Theme.SetByName(themeName);
    }

    /// <summary>
    /// Write the current theme name back into the existing config cell, or create one in the
    /// first empty cell scanning bottom-up / right-to-left so it doesn't clobber user data.
    /// </summary>
    public static void PersistTheme(GridManager grid)
    {
        var loc = Find(grid);
        Dictionary<string, string> map;
        if (loc != null)
        {
            map = CellPrefix.ParseConfig(grid.GetCellValue(loc.Value.row, loc.Value.col));
        }
        else
        {
            loc = FindEmpty(grid);
            if (loc == null) return; // nowhere to put it
            map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
        map["theme"] = Theme.Current.Name;
        grid.SetCellValue(loc.Value.row, loc.Value.col, CellPrefix.FormatConfig(map));
    }

    private static (int row, int col)? Find(GridManager grid)
    {
        for (int r = 0; r < grid.RowCount; r++)
            for (int c = 0; c < grid.ColumnCount; c++)
                if (CellPrefix.IsConfig(grid.GetCellValue(r, c)))
                    return (r, c);
        return null;
    }

    private static (int row, int col)? FindEmpty(GridManager grid)
    {
        for (int r = grid.RowCount - 1; r >= 0; r--)
            for (int c = grid.ColumnCount - 1; c >= 0; c--)
                if (string.IsNullOrEmpty(grid.GetCellValue(r, c)))
                    return (r, c);
        return null;
    }
}
