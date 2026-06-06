namespace ExcelConsole;

public class GridManager
{
    private readonly string[,] _data;
    private readonly bool[,] _isFile;
    private readonly string[,] _filePath;
    public int ColumnCount { get; }
    public int RowCount { get; }
    public bool IsDirty { get; private set; }

    // Extension output overlay — values visible on grid but NOT persisted to CSV.
    // Key: (row, col). Extensions write here via SetExtensionCellValue().
    private readonly Dictionary<(int row, int col), string> _extensionOverlay = new();

    // ── Undo / Redo ────────────────────────────────────────────────
    private readonly UndoManager _undo = new();
    public void BeginUndoGroup() => _undo.BeginGroup();
    public void EndUndoGroup() => _undo.EndGroup();

    // ── Cursor state ─────────────────────────────────────────────────

    private int _selectedRow;
    private int _selectedCol;

    public (int row, int col) GetCurrentCell() => (_selectedRow, _selectedCol);

    public void SelectCell(int row, int col)
    {
        _selectedRow = Math.Clamp(row, 0, RowCount - 1);
        _selectedCol = Math.Clamp(col, 0, ColumnCount - 1);
    }

    public void MoveUp() { if (_selectedRow > 0) _selectedRow--; }
    public void MoveDown() { if (_selectedRow < RowCount - 1) _selectedRow++; }
    public void MoveLeft() { if (_selectedCol > 0) _selectedCol--; }
    public void MoveRight() { if (_selectedCol < ColumnCount - 1) _selectedCol++; }

    public void ClampSelection()
    {
        _selectedRow = Math.Min(_selectedRow, RowCount - 1);
        _selectedCol = Math.Min(_selectedCol, ColumnCount - 1);
    }

    public GridManager(int availableWidth, int availableHeight, int columnWidth = 20)
    {
        const int rowHeaderWidth = 4;

        ColumnCount = Math.Max(1, (availableWidth - rowHeaderWidth) / columnWidth);
        RowCount = Math.Max(1, availableHeight);

        _data = new string[RowCount, ColumnCount];
        _isFile = new bool[RowCount, ColumnCount];
        _filePath = new string[RowCount, ColumnCount];
        for (int r = 0; r < RowCount; r++)
            for (int c = 0; c < ColumnCount; c++)
            {
                _data[r, c] = "";
                _filePath[r, c] = "";
            }
    }

    public static string GetColumnName(int index)
    {
        string name = "";
        index++;
        while (index > 0)
        {
            index--;
            name = (char)('A' + index % 26) + name;
            index /= 26;
        }
        return name;
    }

    public string GetCellReference(int row, int col)
    {
        return $"{GetColumnName(col)}{row + 1}";
    }

    public string GetCellValue(int row, int col)
    {
        if (row >= 0 && row < RowCount && col >= 0 && col < ColumnCount)
        {
            // Extension overlay takes priority for display but is never saved to CSV.
            if (_extensionOverlay.TryGetValue((row, col), out string? overlayVal))
                return overlayVal;
            return _data[row, col];
        }
        return "";
    }

    /// <summary>
    /// Writes a value that should be visible on screen but NOT persisted to CSV.
    /// Used by extensions so their output doesn't pollute the saved data file.
    /// </summary>
    public void SetExtensionCellValue(int row, int col, string value)
    {
        if (row >= 0 && row < RowCount && col >= 0 && col < ColumnCount)
        {
            if (string.IsNullOrEmpty(value))
                _extensionOverlay.Remove((row, col));
            else
                _extensionOverlay[(row, col)] = value;
        }
    }

    /// <summary>
    /// Clears all extension overlay values in the rectangular region starting at
    /// (anchorRow, anchorCol) with the given height and width. Call this when an
    /// extension reactivates so stale output doesn't linger.
    /// </summary>
    public void ClearExtensionOverlay(int anchorRow, int anchorCol, int height, int width)
    {
        for (int r = anchorRow; r < anchorRow + height; r++)
            for (int c = anchorCol; c < anchorCol + width; c++)
                _extensionOverlay.Remove((r, c));
    }

    public string GetSelectedCellValue() => GetCellValue(_selectedRow, _selectedCol);
    public void SetSelectedCellValue(string value) => SetCellValue(_selectedRow, _selectedCol, value);
    public void AppendToSelectedCell(char ch) { var cur = GetSelectedCellValue(); SetSelectedCellValue(cur + ch); }

    public void SetCellValue(int row, int col, string value)
    {
        if (row >= 0 && row < RowCount && col >= 0 && col < ColumnCount)
        {
            _undo.RecordChange(row, col, _data[row, col], value);
            _data[row, col] = value;
            IsDirty = true;
        }
    }

    /// <summary>Apply cell values directly without recording undo (used by Undo/Redo restore).</summary>
    private void SetCellValueRaw(int row, int col, string value)
    {
        if (row >= 0 && row < RowCount && col >= 0 && col < ColumnCount)
        {
            _data[row, col] = value;
            IsDirty = true;
        }
    }

    public bool Undo()
    {
        var changes = _undo.Undo();
        if (changes == null) return false;
        foreach (var (row, col, value) in changes)
            SetCellValueRaw(row, col, value);
        return true;
    }

    public bool Redo()
    {
        var changes = _undo.Redo();
        if (changes == null) return false;
        foreach (var (row, col, value) in changes)
            SetCellValueRaw(row, col, value);
        return true;
    }

    public void SetFileEntry(int row, int col, string displayName, string fullPath)
    {
        if (row < 0 || row >= RowCount || col < 0 || col >= ColumnCount) return;
        _data[row, col] = displayName;
        _isFile[row, col] = true;
        _filePath[row, col] = fullPath;
    }

    public bool IsFileEntry(int row, int col)
    {
        if (row < 0 || row >= RowCount || col < 0 || col >= ColumnCount) return false;
        return _isFile[row, col];
    }

    public string GetFilePath(int row, int col)
    {
        if (row < 0 || row >= RowCount || col < 0 || col >= ColumnCount) return "";
        return _filePath[row, col];
    }

    public void ClearRow(int row)
    {
        if (row < 0 || row >= RowCount) return;
        _undo.BeginGroup();
        for (int c = 0; c < ColumnCount; c++)
        {
            _undo.RecordChange(row, c, _data[row, c], "");
            _data[row, c] = "";
        }
        _undo.EndGroup();
        IsDirty = true;
    }

    public void DeleteRow(int row)
    {
        if (row < 0 || row >= RowCount) return;
        _undo.BeginGroup();
        for (int r = row; r < RowCount - 1; r++)
            for (int c = 0; c < ColumnCount; c++)
            {
                _undo.RecordChange(r, c, _data[r, c], _data[r + 1, c]);
                _data[r, c] = _data[r + 1, c];
            }
        for (int c = 0; c < ColumnCount; c++)
        {
            _undo.RecordChange(RowCount - 1, c, _data[RowCount - 1, c], "");
            _data[RowCount - 1, c] = "";
        }
        _undo.EndGroup();
        IsDirty = true;
    }

    public void DeleteSelectedRow()
    {
        DeleteRow(_selectedRow);
        ClampSelection();
    }

    public void ShiftRowsDown(int fromRow)
    {
        if (fromRow < 0 || fromRow >= RowCount) return;
        _undo.BeginGroup();
        for (int r = RowCount - 1; r > fromRow; r--)
            for (int c = 0; c < ColumnCount; c++)
            {
                _undo.RecordChange(r, c, _data[r, c], _data[r - 1, c]);
                _data[r, c] = _data[r - 1, c];
            }
        for (int c = 0; c < ColumnCount; c++)
        {
            _undo.RecordChange(fromRow, c, _data[fromRow, c], "");
            _data[fromRow, c] = "";
        }
        _undo.EndGroup();
        IsDirty = true;
    }

    public void ShiftRowsUp(int fromRow)
    {
        if (fromRow < 0 || fromRow >= RowCount) return;
        _undo.BeginGroup();
        for (int r = fromRow; r < RowCount - 1; r++)
            for (int c = 0; c < ColumnCount; c++)
            {
                _undo.RecordChange(r, c, _data[r, c], _data[r + 1, c]);
                _data[r, c] = _data[r + 1, c];
            }
        for (int c = 0; c < ColumnCount; c++)
        {
            _undo.RecordChange(RowCount - 1, c, _data[RowCount - 1, c], "");
            _data[RowCount - 1, c] = "";
        }
        _undo.EndGroup();
        IsDirty = true;
    }

    public void ShiftSelectedRowDown() => ShiftRowsDown(_selectedRow);
    public void ShiftSelectedRowUp() => ShiftRowsUp(_selectedRow);

    /// <summary>
    /// Duplicates the given row by inserting a copy immediately below it.
    /// The original row is preserved; all rows from (row+1) down shift down by one.
    /// </summary>
    public void DuplicateRow(int row)
    {
        if (row < 0 || row >= RowCount) return;
        _undo.BeginGroup();
        // Shift everything from (row+1) downward to make room
        for (int r = RowCount - 1; r > row + 1; r--)
            for (int c = 0; c < ColumnCount; c++)
            {
                _undo.RecordChange(r, c, _data[r, c], _data[r - 1, c]);
                _data[r, c] = _data[r - 1, c];
            }
        // Write the copy into (row+1)
        for (int c = 0; c < ColumnCount; c++)
        {
            string orig = _data[row, c];
            _undo.RecordChange(row + 1, c, _data[row + 1, c], orig);
            _data[row + 1, c] = orig;
        }
        _undo.EndGroup();
        IsDirty = true;
    }



    private int _lastSortCol = -1;
    private bool _lastSortAscending = true;

    /// <summary>
    /// Sort all rows by the given column. Toggles ascending/descending on
    /// repeated calls for the same column. Numeric values sort numerically;
    /// non-numeric values sort lexicographically. Empty cells sort last.
    /// </summary>
    public void SortByColumn(int col)
    {
        if (col < 0 || col >= ColumnCount) return;

        // Toggle direction when sorting same column again
        bool ascending;
        if (col == _lastSortCol)
        {
            ascending = !_lastSortAscending;
        }
        else
        {
            ascending = true;
        }
        _lastSortCol = col;
        _lastSortAscending = ascending;

        // Find last non-empty row to avoid sorting trailing blank rows
        int lastDataRow = -1;
        for (int r = RowCount - 1; r >= 0; r--)
        {
            for (int c = 0; c < ColumnCount; c++)
            {
                if (!string.IsNullOrEmpty(_data[r, c]))
                {
                    lastDataRow = r;
                    break;
                }
            }
            if (lastDataRow >= 0) break;
        }
        if (lastDataRow <= 0) return; // Nothing to sort (0 or 1 data rows)

        int sortCount = lastDataRow + 1;

        // Build row indices for sorting
        var indices = Enumerable.Range(0, sortCount).ToArray();

        Array.Sort(indices, (a, b) =>
        {
            string va = _data[a, col];
            string vb = _data[b, col];
            bool emptyA = string.IsNullOrEmpty(va);
            bool emptyB = string.IsNullOrEmpty(vb);
            if (emptyA && emptyB) return 0;
            if (emptyA) return 1;  // empties last regardless of direction
            if (emptyB) return -1;

            bool numA = double.TryParse(va, out double da);
            bool numB = double.TryParse(vb, out double db);

            int cmp;
            if (numA && numB)
                cmp = da.CompareTo(db);
            else
                cmp = string.Compare(va, vb, StringComparison.OrdinalIgnoreCase);

            return ascending ? cmp : -cmp;
        });

        // Record undo for all cells that change
        _undo.BeginGroup();

        // Build sorted copy of data rows
        var sortedData = new string[sortCount, ColumnCount];
        for (int r = 0; r < sortCount; r++)
            for (int c = 0; c < ColumnCount; c++)
                sortedData[r, c] = _data[indices[r], c];

        // Apply and record changes
        for (int r = 0; r < sortCount; r++)
            for (int c = 0; c < ColumnCount; c++)
            {
                if (_data[r, c] != sortedData[r, c])
                {
                    _undo.RecordChange(r, c, _data[r, c], sortedData[r, c]);
                    _data[r, c] = sortedData[r, c];
                }
            }

        _undo.EndGroup();
        IsDirty = true;
    }

    public double? GetColumnSum(int col)
    {
        if (col < 0 || col >= ColumnCount)
            return null;

        double sum = 0;
        bool hasNumber = false;

        for (int r = 0; r < RowCount; r++)
        {
            if (double.TryParse(_data[r, col], System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double num))
            {
                sum += num;
                hasNumber = true;
            }
        }

        return hasNumber ? sum : null;
    }

    public double? GetRowProduct(int row)
    {
        if (row < 0 || row >= RowCount)
            return null;

        double product = 1;
        bool hasNumber = false;

        for (int c = 0; c < ColumnCount; c++)
        {
            if (double.TryParse(_data[row, c], System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double num))
            {
                product *= num;
                hasNumber = true;
            }
        }

        return hasNumber ? product : null;
    }

    public void SaveToCsv(string path)
    {
        // Read existing file to preserve rows/columns beyond the grid bounds
        List<List<string>>? existingRows = null;
        int existingRowCount = 0;
        int existingMaxCols = 0;
        if (File.Exists(path))
        {
            try
            {
                var lines = File.ReadAllLines(path);
                existingRows = new List<List<string>>(lines.Length);
                foreach (var line in lines)
                {
                    var parsed = ParseCsvLine(line);
                    existingRows.Add(parsed);
                    if (parsed.Count > existingMaxCols)
                        existingMaxCols = parsed.Count;
                }
                existingRowCount = existingRows.Count;
            }
            catch { /* If we can't read, just save what we have */ }
        }

        int totalRows = Math.Max(RowCount, existingRowCount);
        int totalCols = Math.Max(ColumnCount, existingMaxCols);

        using var writer = new StreamWriter(path);
        for (int r = 0; r < totalRows; r++)
        {
            var fields = new string[totalCols];
            for (int c = 0; c < totalCols; c++)
            {
                if (r < RowCount && c < ColumnCount)
                {
                    // Within grid bounds: use grid data
                    fields[c] = EscapeCsvField(_isFile[r, c] ? "" : _data[r, c]);
                }
                else if (existingRows != null && r < existingRows.Count && c < existingRows[r].Count)
                {
                    // Beyond grid bounds: preserve existing file data
                    fields[c] = EscapeCsvField(existingRows[r][c]);
                }
                else
                {
                    fields[c] = "";
                }
            }
            writer.WriteLine(string.Join(",", fields));
        }
        IsDirty = false;
    }

    /// <summary>
    /// Writes the grid as a GitHub-flavored Markdown table to <paramref name="path"/>.
    /// First row of the grid is treated as the header.
    /// Trailing rows and columns that are entirely empty are trimmed.
    /// </summary>
    public void SaveToMarkdown(string path)
    {
        using var writer = new StreamWriter(path);
        WriteMarkdownTo(writer);
    }

    /// <summary>
    /// Writes the grid as a GitHub-flavored Markdown table to an arbitrary <see cref="TextWriter"/>.
    /// Used by `--export-md -` to write to stdout for piping.
    /// </summary>
    public void WriteMarkdownTo(TextWriter writer)
    {
        // Find the bottom-most non-empty row and right-most non-empty column.
        int lastRow = -1, lastCol = -1;
        for (int r = 0; r < RowCount; r++)
            for (int c = 0; c < ColumnCount; c++)
                if (!string.IsNullOrEmpty(_data[r, c]) && !_isFile[r, c])
                {
                    if (r > lastRow) lastRow = r;
                    if (c > lastCol) lastCol = c;
                }
        if (lastRow < 0 || lastCol < 0) return;

        for (int r = 0; r <= lastRow; r++)
        {
            writer.Write('|');
            for (int c = 0; c <= lastCol; c++)
            {
                string v = _isFile[r, c] ? "" : (_data[r, c] ?? "");
                v = v.Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
                writer.Write(' ');
                writer.Write(v);
                writer.Write(" |");
            }
            writer.WriteLine();
            if (r == 0)
            {
                writer.Write('|');
                for (int c = 0; c <= lastCol; c++) writer.Write("---|");
                writer.WriteLine();
            }
        }
    }

    public void SaveToHtml(string path)
    {
        using var writer = new StreamWriter(path);
        WriteHtmlTo(writer);
    }

    /// <summary>
    /// Writes the grid as a self-contained HTML table with inline styling.
    /// Used by `--export-html -` to write to stdout for piping.
    /// </summary>
    public void WriteHtmlTo(TextWriter writer)
    {
        int lastRow = -1, lastCol = -1;
        for (int r = 0; r < RowCount; r++)
            for (int c = 0; c < ColumnCount; c++)
                if (!string.IsNullOrEmpty(_data[r, c]) && !_isFile[r, c])
                {
                    if (r > lastRow) lastRow = r;
                    if (c > lastCol) lastCol = c;
                }
        if (lastRow < 0 || lastCol < 0) return;

        writer.WriteLine("<!DOCTYPE html>");
        writer.WriteLine("<html lang=\"en\"><head><meta charset=\"utf-8\">");
        writer.WriteLine("<title>QuickSheet Export</title>");
        writer.WriteLine("<style>");
        writer.WriteLine("body{font-family:system-ui,-apple-system,sans-serif;background:#0d1117;color:#e6edf3;padding:2rem}");
        writer.WriteLine("table{border-collapse:collapse;width:100%}");
        writer.WriteLine("th,td{border:1px solid #30363d;padding:6px 10px;text-align:left}");
        writer.WriteLine("th{background:#161b22;color:#8b949e;font-weight:600}");
        writer.WriteLine("tr:nth-child(even) td{background:#161b22}");
        writer.WriteLine("tr:hover td{background:#1c2129}");
        writer.WriteLine("a{color:#58a6ff}");
        writer.WriteLine(".num{text-align:right;font-variant-numeric:tabular-nums}");
        writer.WriteLine("</style></head><body>");
        writer.WriteLine("<table>");

        for (int r = 0; r <= lastRow; r++)
        {
            string tag = r == 0 ? "th" : "td";
            writer.Write("<tr>");
            for (int c = 0; c <= lastCol; c++)
            {
                string v = _isFile[r, c] ? "" : (_data[r, c] ?? "");
                string escaped = System.Net.WebUtility.HtmlEncode(v);
                string cls = "";

                // Auto-detect URLs and make them clickable
                if (v.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    v.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    escaped = $"<a href=\"{System.Net.WebUtility.HtmlEncode(v)}\">{escaped}</a>";
                }

                // Detect numeric cells for right-alignment
                if (r > 0 && double.TryParse(v, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out _))
                {
                    cls = " class=\"num\"";
                }

                writer.Write($"<{tag}{cls}>{escaped}</{tag}>");
            }
            writer.WriteLine("</tr>");
        }

        writer.WriteLine("</table>");
        writer.WriteLine("<p style=\"color:#8b949e;font-size:0.8rem;margin-top:1rem\">Exported by <a href=\"https://github.com/cemheren/QuickSheet\">QuickSheet</a></p>");
        writer.WriteLine("</body></html>");
    }

    public void SaveToJson(string path)
    {
        using var writer = new StreamWriter(path);
        WriteJsonTo(writer);
    }

    /// <summary>
    /// Writes the grid as a JSON array of objects (first row = keys).
    /// Used by `--export-json -` to write to stdout for piping.
    /// </summary>
    public void WriteJsonTo(TextWriter writer)
    {
        int lastRow = -1, lastCol = -1;
        for (int r = 0; r < RowCount; r++)
            for (int c = 0; c < ColumnCount; c++)
                if (!string.IsNullOrEmpty(_data[r, c]) && !_isFile[r, c])
                {
                    if (r > lastRow) lastRow = r;
                    if (c > lastCol) lastCol = c;
                }
        if (lastRow < 0 || lastCol < 0) { writer.WriteLine("[]"); return; }

        // First row = column headers (keys)
        var keys = new string[lastCol + 1];
        for (int c = 0; c <= lastCol; c++)
        {
            string v = _isFile[0, c] ? "" : (_data[0, c] ?? "");
            keys[c] = v.Length > 0 ? v : $"col{c}";
        }

        writer.WriteLine("[");
        for (int r = 1; r <= lastRow; r++)
        {
            writer.Write("  {");
            for (int c = 0; c <= lastCol; c++)
            {
                string v = _isFile[r, c] ? "" : (_data[r, c] ?? "");
                string key = JsonEscape(keys[c]);
                string val = JsonEscape(v);

                // Try to output numbers without quotes
                if (double.TryParse(v, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double num))
                {
                    writer.Write($"\"{key}\":{num}");
                }
                else
                {
                    writer.Write($"\"{key}\":\"{val}\"");
                }

                if (c < lastCol) writer.Write(",");
            }
            writer.Write("}");
            if (r < lastRow) writer.Write(",");
            writer.WriteLine();
        }
        writer.WriteLine("]");
    }

    private static string JsonEscape(string s)
    {
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
    }

    public void LoadFromCsv(string path)
    {
        if (!File.Exists(path)) return;
        var lines = File.ReadAllLines(path);
        for (int r = 0; r < Math.Min(lines.Length, RowCount); r++)
        {
            var fields = ParseCsvLine(lines[r]);
            for (int c = 0; c < Math.Min(fields.Count, ColumnCount); c++)
                _data[r, c] = fields[c];
        }
    }

    /// <summary>
    /// Re-reads the CSV and merges external changes into the grid.
    /// Empty local cells are overwritten. Conflicts produce a "c: " marker.
    /// File-entry cells are skipped.
    /// </summary>
    public bool MergeFromCsv(string path)
    {
        if (!File.Exists(path)) return false;
        string[] lines;
        try { lines = File.ReadAllLines(path); }
        catch { return false; }

        bool changed = false;
        for (int r = 0; r < Math.Min(lines.Length, RowCount); r++)
        {
            var fields = ParseCsvLine(lines[r]);
            for (int c = 0; c < Math.Min(fields.Count, ColumnCount); c++)
            {
                if (_isFile[r, c]) continue;

                string external = fields[c];
                string local = _data[r, c];

                if (local == external) continue;

                if (string.IsNullOrEmpty(local))
                {
                    _data[r, c] = external;
                    changed = true;
                }
                else if (!string.IsNullOrEmpty(external))
                {
                    _data[r, c] = $"c: {external}({local})";
                    changed = true;
                }
            }
        }
        return changed;
    }

    private static string EscapeCsvField(string field)
    {
        if (field.Contains(',') || field.Contains('"') || field.Contains('\n'))
            return "\"" + field.Replace("\"", "\"\"") + "\"";
        return field;
    }

    internal static List<string> ParseCsvLine(string line)
    {
        var fields = new List<string>();
        int i = 0;
        while (i <= line.Length)
        {
            if (i == line.Length) { fields.Add(""); break; }

            if (line[i] == '"')
            {
                // Quoted field
                i++;
                var field = new System.Text.StringBuilder();
                while (i < line.Length)
                {
                    if (line[i] == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            field.Append('"');
                            i += 2;
                        }
                        else
                        {
                            i++; // closing quote
                            break;
                        }
                    }
                    else
                    {
                        field.Append(line[i]);
                        i++;
                    }
                }
                fields.Add(field.ToString());
                if (i < line.Length && line[i] == ',') i++; // skip comma
            }
            else
            {
                // Unquoted field
                int start = i;
                while (i < line.Length && line[i] != ',') i++;
                fields.Add(line[start..i]);
                if (i < line.Length) i++; // skip comma
            }
        }
        return fields;
    }

    // ── Inline resolution ────────────────────────────────────────────

    /// <summary>
    /// Resolves an inline reference chain starting from (row, col).
    /// Returns the final resolved cell value, or null if the cell isn't an inline ref.
    /// Uses a visited set for cycle detection.
    /// </summary>
    public string? ResolveInline(int row, int col, HashSet<(int, int)>? visited = null)
    {
        string value = GetCellValue(row, col);
        if (!CellPrefix.IsInline(value)) return null;

        visited ??= new HashSet<(int, int)>();
        if (!visited.Add((row, col))) return "[circular]";
        if (visited.Count > 7) return "[too deep]";

        // Expand {A1::C10} refs in the i: value before parsing the target cell
        string expanded = CellPrefix.ExpandCellReferences(value, this);
        var target = CellPrefix.ParseInlineRef(expanded);
        if (target == null) return "[invalid ref]";

        int tr = target.Value.row, tc = target.Value.col;
        if (tr < 0 || tr >= RowCount || tc < 0 || tc >= ColumnCount)
            return "[out of bounds]";

        string targetValue = GetCellValue(tr, tc);

        // If target is also inline, resolve recursively
        if (CellPrefix.IsInline(targetValue))
            return ResolveInline(tr, tc, visited);

        return targetValue;
    }

    public string? ResolveSelectedInline(HashSet<(int, int)>? visited = null)
        => ResolveInline(_selectedRow, _selectedCol, visited);

    /// <summary>
    /// Gets the display value for a cell, resolving inline refs if applicable.
    /// </summary>
    public string GetDisplayValue(int row, int col, HashSet<(int, int)>? visited = null)
    {
        string raw = GetCellValue(row, col);

        if (CellPrefix.IsInline(raw))
        {
            visited ??= new HashSet<(int, int)>();
            return ResolveInline(row, col, visited) ?? raw;
        }

        return raw;
    }
}
