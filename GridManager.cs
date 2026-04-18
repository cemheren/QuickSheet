namespace ExcelConsole;

public class GridManager
{
    private readonly string[,] _data;
    private readonly bool[,] _isFile;
    private readonly string[,] _filePath;
    public int ColumnCount { get; }
    public int RowCount { get; }
    public bool IsDirty { get; private set; }

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
            return _data[row, col];
        return "";
    }

    public string GetSelectedCellValue() => GetCellValue(_selectedRow, _selectedCol);
    public void SetSelectedCellValue(string value) => SetCellValue(_selectedRow, _selectedCol, value);
    public void AppendToSelectedCell(char ch) { var cur = GetSelectedCellValue(); SetSelectedCellValue(cur + ch); }

    public void SetCellValue(int row, int col, string value)
    {
        if (row >= 0 && row < RowCount && col >= 0 && col < ColumnCount)
        {
            _data[row, col] = value;
            IsDirty = true;
        }
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
        for (int c = 0; c < ColumnCount; c++)
            _data[row, c] = "";
        IsDirty = true;
    }

    public void DeleteRow(int row)
    {
        if (row < 0 || row >= RowCount) return;
        for (int r = row; r < RowCount - 1; r++)
            for (int c = 0; c < ColumnCount; c++)
                _data[r, c] = _data[r + 1, c];
        for (int c = 0; c < ColumnCount; c++)
            _data[RowCount - 1, c] = "";
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
        for (int r = RowCount - 1; r > fromRow; r--)
            for (int c = 0; c < ColumnCount; c++)
                _data[r, c] = _data[r - 1, c];
        for (int c = 0; c < ColumnCount; c++)
            _data[fromRow, c] = "";
        IsDirty = true;
    }

    public void ShiftRowsUp(int fromRow)
    {
        if (fromRow < 0 || fromRow >= RowCount) return;
        for (int r = fromRow; r < RowCount - 1; r++)
            for (int c = 0; c < ColumnCount; c++)
                _data[r, c] = _data[r + 1, c];
        for (int c = 0; c < ColumnCount; c++)
            _data[RowCount - 1, c] = "";
        IsDirty = true;
    }

    public void ShiftSelectedRowDown() => ShiftRowsDown(_selectedRow);
    public void ShiftSelectedRowUp() => ShiftRowsUp(_selectedRow);

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

    private static List<string> ParseCsvLine(string line)
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
