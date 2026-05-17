using System.Globalization;
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

    public static bool IsSparkline(string value) =>
        value.StartsWith("s: ", StringComparison.OrdinalIgnoreCase);

    private static readonly char[] SparklineChars = ['▁', '▂', '▃', '▄', '▅', '▆', '▇', '█'];

    /// <summary>
    /// Renders an "s: 1,2,3,4,5" or "s: A1::A10" cell as unicode block-bar sparkline characters.
    /// The range form pulls numeric values from <paramref name="grid"/> when supplied.
    /// Returns the rendered glyph string, or null if the value is not a sparkline cell or fails to parse.
    /// </summary>
    public static string? RenderSparkline(string value, GridManager? grid = null)
    {
        if (!IsSparkline(value)) return null;
        string rest = value[3..].Trim();
        if (rest.Length == 0) return null;

        List<double> nums;

        // Range form: `s: A1::A10` or `s: A1-A10`. Only honored when a grid is supplied.
        var sparkRange = ParseCellRange(rest);
        if (sparkRange != null && grid != null)
        {
            var (r1, c1, r2, c2) = sparkRange.Value;
            r2 = Math.Min(r2, grid.RowCount - 1);
            c2 = Math.Min(c2, grid.ColumnCount - 1);
            if (r1 >= grid.RowCount || c1 >= grid.ColumnCount) return null;

            nums = new List<double>();
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                {
                    string v = grid.GetCellValue(r, c);
                    if (double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double n))
                        nums.Add(n);
                }
            if (nums.Count == 0) return null;
        }
        else
        {
            string[] parts = rest.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return null;

            nums = new List<double>(parts.Length);
            foreach (string p in parts)
            {
                if (!double.TryParse(p, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double n))
                    return null;
                nums.Add(n);
            }
        }

        double min = nums[0], max = nums[0];
        foreach (double n in nums) { if (n < min) min = n; if (n > max) max = n; }
        double range = max - min;

        var sb = new System.Text.StringBuilder(nums.Count);
        foreach (double n in nums)
        {
            int idx = range > 0
                ? (int)Math.Round((n - min) / range * (SparklineChars.Length - 1))
                : SparklineChars.Length / 2;
            if (idx < 0) idx = 0;
            if (idx >= SparklineChars.Length) idx = SparklineChars.Length - 1;
            sb.Append(SparklineChars[idx]);
        }
        return sb.ToString();
    }

    // ── Color prefix ───────────────────────────────────────────────

    private static readonly Dictionary<string, ConsoleColor> ColorMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["red"] = ConsoleColor.DarkRed,
        ["green"] = ConsoleColor.DarkGreen,
        ["blue"] = ConsoleColor.DarkBlue,
        ["yellow"] = ConsoleColor.DarkYellow,
        ["cyan"] = ConsoleColor.DarkCyan,
        ["magenta"] = ConsoleColor.DarkMagenta,
        ["white"] = ConsoleColor.White,
        ["gray"] = ConsoleColor.DarkGray,
        ["grey"] = ConsoleColor.DarkGray,
    };

    /// <summary>
    /// Checks if a cell value uses the color prefix: "c:color: text"
    /// or the value-driven color prefix: "c?: rule, rule, *=default: value".
    /// </summary>
    public static bool IsColored(string value) =>
        (value.StartsWith("c:", StringComparison.OrdinalIgnoreCase) ||
         value.StartsWith("c?:", StringComparison.OrdinalIgnoreCase))
        && ParseColor(value) != null;

    /// <summary>
    /// Parses "c:red: some text" into (ConsoleColor, displayText), or
    /// "c?: &gt;100=red, &gt;50=yellow, *=green: 73" into the colour the value matches.
    /// Returns null if the prefix or rules don't yield a colour.
    /// </summary>
    public static (ConsoleColor bg, string text)? ParseColor(string value)
    {
        if (value.StartsWith("c?:", StringComparison.OrdinalIgnoreCase))
            return ParseConditionalColor(value);
        if (!value.StartsWith("c:", StringComparison.OrdinalIgnoreCase)) return null;
        // Format: c:COLOR: text
        int secondColon = value.IndexOf(':', 2);
        if (secondColon < 0) return null;

        string colorName = value[2..secondColon].Trim();
        if (!ColorMap.TryGetValue(colorName, out var color)) return null;

        string text = value[(secondColon + 1)..].TrimStart();
        return (color, text);
    }

    /// <summary>
    /// Parses "c?: &gt;100=red, &gt;50=yellow, *=green: 73". Rules are comma-separated
    /// "OP NUMBER = COLOR" where OP is one of &gt; &lt; &gt;= &lt;= =, plus a "*=COLOR"
    /// default. Rules are evaluated left-to-right; first match wins. Non-numeric values
    /// only match the "*=COLOR" default. Returns null if no rule matches.
    /// </summary>
    static (ConsoleColor bg, string text)? ParseConditionalColor(string value)
    {
        int lastColon = value.LastIndexOf(':');
        if (lastColon <= 2) return null;

        string rulesPart = value[3..lastColon];
        string text = value[(lastColon + 1)..].TrimStart();

        bool numeric = double.TryParse(text.Trim(), NumberStyles.Float,
            CultureInfo.InvariantCulture, out double v);

        foreach (var rawRule in rulesPart.Split(','))
        {
            var rule = rawRule.Trim();
            if (rule.Length == 0) continue;
            int eq = rule.IndexOf('=');
            if (eq <= 0) continue;
            string condition = rule[..eq].Trim();
            string colorName = rule[(eq + 1)..].Trim();
            if (!ColorMap.TryGetValue(colorName, out var color)) continue;

            if (condition == "*") return (color, text);
            if (!numeric) continue;

            string op;
            string numStr;
            if (condition.StartsWith(">=")) { op = ">="; numStr = condition[2..]; }
            else if (condition.StartsWith("<=")) { op = "<="; numStr = condition[2..]; }
            else if (condition.StartsWith(">")) { op = ">"; numStr = condition[1..]; }
            else if (condition.StartsWith("<")) { op = "<"; numStr = condition[1..]; }
            else if (condition.StartsWith("=")) { op = "="; numStr = condition[1..]; }
            else continue;

            if (!double.TryParse(numStr.Trim(), NumberStyles.Float,
                CultureInfo.InvariantCulture, out double n)) continue;

            bool match = op switch
            {
                ">" => v > n,
                "<" => v < n,
                ">=" => v >= n,
                "<=" => v <= n,
                "=" => v == n,
                _ => false,
            };
            if (match) return (color, text);
        }
        return null;
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
        // Strip status suffixes that may have been written by older versions
        foreach (var suffix in new[] { "[install failed]", "[bad manifest]", "[start failed]" })
        {
            if (source.EndsWith(suffix))
            {
                source = source[..^suffix.Length].Trim();
                break;
            }
        }
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

        // Default output area: 1 column, 10 rows — enough for most extensions
        const int defaultGridCols = 1;
        const int defaultGridRows = 10;

        // Try to parse trailing gridCols,gridRows if at least 2 parts exist
        if (parts.Length >= 2 &&
            int.TryParse(parts[^2], out int gridCols) &&
            int.TryParse(parts[^1], out int gridRows) &&
            gridCols >= 1 && gridRows >= 1)
        {
            // Last two parts are valid dimensions
            string[] extParams = parts.Length > 2 ? parts[..^2] : [];
            return (prefix, extParams, gridCols, gridRows);
        }

        // No valid trailing dimensions — treat all parts as extension params, use defaults
        return (prefix, parts, defaultGridCols, defaultGridRows);
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
