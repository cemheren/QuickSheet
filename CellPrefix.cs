using System.Text.RegularExpressions;

namespace ExcelConsole;

/// <summary>
/// Parsing utilities for cell prefixes (w:, i:, r:) and cell reference syntax {A1::C10}.
/// </summary>
public static class CellPrefix
{
    // ── Prefix detection ─────────────────────────────────────────────

    public static bool IsWindow(string value) =>
        value.StartsWith("w: ", StringComparison.OrdinalIgnoreCase);

    public static bool IsInline(string value) =>
        value.StartsWith("i: ", StringComparison.OrdinalIgnoreCase);

    public static bool IsCommand(string value) =>
        value.StartsWith("r: ", StringComparison.OrdinalIgnoreCase);

    public static bool IsHyperlink(string value) =>
        value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        value.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

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

    // ── Window prefix parsing ────────────────────────────────────────

    /// <summary>
    /// Parses a "w: F40-H50 optional content" cell value.
    /// Returns the range and the content after the range (may be empty).
    /// </summary>
    public static (int startRow, int startCol, int endRow, int endCol, string content)? ParseWindowDef(string cellValue)
    {
        if (!IsWindow(cellValue)) return null;
        string afterPrefix = cellValue[3..].Trim(); // skip "w: "

        // First token is the range, rest is content
        int spaceIdx = afterPrefix.IndexOf(' ');
        string rangeStr = spaceIdx >= 0 ? afterPrefix[..spaceIdx] : afterPrefix;
        string content = spaceIdx >= 0 ? afterPrefix[(spaceIdx + 1)..] : "";

        var range = ParseCellRange(rangeStr);
        if (range == null) return null;

        return (range.Value.startRow, range.Value.startCol, range.Value.endRow, range.Value.endCol, content);
    }

    /// <summary>
    /// Returns the length of the "w: RANGE " prefix (including trailing space).
    /// Used to protect the prefix during backspace. Returns cellValue.Length if no content.
    /// </summary>
    public static int GetWindowPrefixLength(string cellValue)
    {
        if (!IsWindow(cellValue)) return 0;
        string afterW = cellValue[3..].TrimStart();
        int rangeStart = cellValue.Length - afterW.Length; // position where range starts
        int spaceAfterRange = afterW.IndexOf(' ');
        if (spaceAfterRange < 0) return cellValue.Length; // no content, entire string is prefix
        return rangeStart + spaceAfterRange + 1; // include space after range
    }

    /// <summary>
    /// Parses an "i: A10" cell value. Returns the source cell reference.
    /// </summary>
    public static (int row, int col)? ParseInlineRef(string cellValue)
    {
        if (!IsInline(cellValue)) return null;
        string afterPrefix = cellValue[3..].Trim(); // skip "i: "
        return ParseCellRef(afterPrefix);
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
