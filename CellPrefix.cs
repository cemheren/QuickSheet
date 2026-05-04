using System.Text.RegularExpressions;

namespace ExcelConsole;

/// <summary>
/// Parsing utilities for cell prefixes (i:, r:) and cell reference syntax {A1::C10}.
/// </summary>
public static class CellPrefix
{
    // ── Prefix detection ─────────────────────────────────────────────

    public static bool IsInline(string value) =>
        value.StartsWith("i: ", StringComparison.OrdinalIgnoreCase);

    public static bool IsCommand(string value) =>
        value.StartsWith("r: ", StringComparison.OrdinalIgnoreCase);

    public static bool IsHyperlink(string value) =>
        value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        value.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

    public static bool IsLoop(string value) =>
        value.StartsWith("L: ", StringComparison.Ordinal) ||
        value.StartsWith("l: ", StringComparison.Ordinal);

    /// <summary>
    /// Parses "L: A10,15m" into the target cell (row, col) and interval in minutes.
    /// Returns null if parsing fails.
    /// </summary>
    public static (int row, int col, int minutes)? ParseLoop(string value)
    {
        if (!IsLoop(value)) return null;
        string rest = value[3..].Trim();
        // Expected format: <cellRef>,<N>m
        int commaIdx = rest.IndexOf(',');
        if (commaIdx <= 0) return null;

        string cellPart = rest[..commaIdx].Trim();
        string intervalPart = rest[(commaIdx + 1)..].Trim();

        var cellRef = ParseCellRef(cellPart);
        if (cellRef == null) return null;

        // Strip trailing 'm' and parse number
        if (intervalPart.EndsWith('m') || intervalPart.EndsWith('M'))
            intervalPart = intervalPart[..^1].Trim();
        if (!int.TryParse(intervalPart, out int minutes) || minutes < 1)
            return null;

        return (cellRef.Value.row, cellRef.Value.col, minutes);
    }

    public static bool IsExtension(string value) =>
        value.StartsWith("ext: ", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Parses an "ext: github:user/repo" cell value into the GitHub reference.
    /// Returns null if not a valid extension cell.
    /// </summary>
    public static string? ParseExtensionSource(string value)
    {
        if (!IsExtension(value)) return null;
        string source = value[5..].Trim();
        return string.IsNullOrEmpty(source) ? null : source;
    }

    /// <summary>
    /// Parses a prefixed cell like "wthr: 98112,2,7" into (prefix, params[], gridCols, gridRows).
    /// Last two comma-separated values are always gridCols and gridRows.
    /// Returns null if parsing fails.
    /// </summary>
    public static (string prefix, string[] extParams, int gridCols, int gridRows)? ParseExtensionCall(
        string value, HashSet<string> knownPrefixes)
    {
        int colonIdx = value.IndexOf(':');
        if (colonIdx <= 0) return null;

        string prefix = value[..colonIdx].Trim().ToLowerInvariant();
        if (!knownPrefixes.Contains(prefix)) return null;

        string rest = value[(colonIdx + 1)..].Trim();
        if (string.IsNullOrEmpty(rest)) return null;

        string[] parts = rest.Split(',', StringSplitOptions.TrimEntries);
        if (parts.Length < 2) return null;

        // Last two are always gridCols, gridRows
        if (!int.TryParse(parts[^2], out int gridCols) || !int.TryParse(parts[^1], out int gridRows))
            return null;
        if (gridCols < 1 || gridRows < 1) return null;

        string[] extParams = parts.Length > 2 ? parts[..^2] : [];
        return (prefix, extParams, gridCols, gridRows);
    }

    // ── Cell reference parsing ───────────────────────────────────────

    /// <summary>
    /// Parses a column letter string (e.g., "A", "AB") into a zero-based column index.
    /// Returns -1 if invalid.
    /// </summary>
    public static int ParseColumnLetters(string letters)
    {
        if (string.IsNullOrEmpty(letters)) return -1;
        int col = 0;
        foreach (char ch in letters.ToUpperInvariant())
        {
            if (ch < 'A' || ch > 'Z') return -1;
            col = col * 26 + (ch - 'A' + 1);
        }
        return col - 1; // zero-based
    }

    /// <summary>
    /// Parses a cell reference like "A1" or "AB123" into (row, col), both zero-based.
    /// Returns null if invalid.
    /// </summary>
    public static (int row, int col)? ParseCellRef(string cellRef)
    {
        if (string.IsNullOrWhiteSpace(cellRef)) return null;
        cellRef = cellRef.Trim().ToUpperInvariant();

        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        if (i == 0 || i == cellRef.Length) return null;

        string letters = cellRef[..i];
        string digits = cellRef[i..];

        int col = ParseColumnLetters(letters);
        if (col < 0) return null;
        if (!int.TryParse(digits, out int rowNum) || rowNum < 1) return null;

        return (rowNum - 1, col); // zero-based
    }

    /// <summary>
    /// Parses a cell range like "F40-H50" or "A1::C10" into start/end (row, col) pairs.
    /// Supports both '-' and '::' as separators.
    /// Returns null if invalid.
    /// </summary>
    public static (int startRow, int startCol, int endRow, int endCol)? ParseCellRange(string rangeStr)
    {
        if (string.IsNullOrWhiteSpace(rangeStr)) return null;
        rangeStr = rangeStr.Trim();

        string[] parts;
        if (rangeStr.Contains("::"))
            parts = rangeStr.Split("::", 2, StringSplitOptions.TrimEntries);
        else if (rangeStr.Contains('-'))
            parts = rangeStr.Split('-', 2, StringSplitOptions.TrimEntries);
        else
            return null;

        if (parts.Length != 2) return null;

        var start = ParseCellRef(parts[0]);
        var end = ParseCellRef(parts[1]);
        if (start == null || end == null) return null;

        // Normalize so start <= end
        int startRow = Math.Min(start.Value.row, end.Value.row);
        int startCol = Math.Min(start.Value.col, end.Value.col);
        int endRow = Math.Max(start.Value.row, end.Value.row);
        int endCol = Math.Max(start.Value.col, end.Value.col);

        return (startRow, startCol, endRow, endCol);
    }

    /// <summary>
    /// Parses an "i: A10" or "i: A10,5,3" cell value.
    /// Returns the source cell reference and optional span dimensions (cols, rows).
    /// Defaults to (1,1) if no dimensions given.
    /// </summary>
    public static (int row, int col, int spanCols, int spanRows)? ParseInlineRef(string cellValue)
    {
        if (!IsInline(cellValue)) return null;
        string afterPrefix = cellValue[3..].Trim();

        string[] parts = afterPrefix.Split(',');
        var cellRef = ParseCellRef(parts[0].Trim());
        if (cellRef == null) return null;

        int spanCols = 1, spanRows = 1;
        if (parts.Length >= 3)
        {
            int.TryParse(parts[1].Trim(), out spanCols);
            int.TryParse(parts[2].Trim(), out spanRows);
            if (spanCols < 1) spanCols = 1;
            if (spanRows < 1) spanRows = 1;
        }

        return (cellRef.Value.row, cellRef.Value.col, spanCols, spanRows);
    }

    // ── Cell reference expansion {A1::C10} ───────────────────────────

    private static readonly Regex CellRangeRefPattern = new(
        @"\{([A-Za-z]+\d+)::([A-Za-z]+\d+)\}",
        RegexOptions.Compiled);

    /// <summary>
    /// Expands all {A1::C10} references in a string by replacing them with
    /// stringified cell data from the grid. Uses raw cell values (no transitive execution).
    /// </summary>
    public static string ExpandCellReferences(string text, GridManager grid, HashSet<(int, int)>? visited = null)
    {
        if (string.IsNullOrEmpty(text) || !text.Contains('{')) return text;

        return CellRangeRefPattern.Replace(text, match =>
        {
            var start = ParseCellRef(match.Groups[1].Value);
            var end = ParseCellRef(match.Groups[2].Value);
            if (start == null || end == null) return match.Value;

            int r1 = Math.Min(start.Value.row, end.Value.row);
            int c1 = Math.Min(start.Value.col, end.Value.col);
            int r2 = Math.Max(start.Value.row, end.Value.row);
            int c2 = Math.Max(start.Value.col, end.Value.col);

            // Clamp to grid bounds
            r2 = Math.Min(r2, grid.RowCount - 1);
            c2 = Math.Min(c2, grid.ColumnCount - 1);
            if (r1 >= grid.RowCount || c1 >= grid.ColumnCount) return "";

            // Check for circular references
            if (visited != null)
            {
                for (int r = r1; r <= r2; r++)
                    for (int c = c1; c <= c2; c++)
                        if (visited.Contains((r, c)))
                            return "[circular]";
            }

            var rows = new List<string>();
            for (int r = r1; r <= r2; r++)
            {
                var cols = new List<string>();
                for (int c = c1; c <= c2; c++)
                {
                    string val = grid.GetCellValue(r, c);
                    // Escape tabs and newlines in values
                    val = val.Replace("\t", "\\t").Replace("\n", "\\n").Replace("\r", "");
                    cols.Add(val);
                }
                rows.Add(string.Join("\t", cols));
            }
            return string.Join("\\n", rows);
        });
    }
}
