namespace ExcelConsole;

public class GridManager
{
    private readonly string[,] _data;
    private readonly bool[,] _isFile;
    private readonly string[,] _filePath;
    public int ColumnCount { get; }
    public int RowCount { get; }

    public GridManager(int availableWidth, int availableHeight)
    {
        const int columnWidth = 20;
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

    public void SetCellValue(int row, int col, string value)
    {
        if (row >= 0 && row < RowCount && col >= 0 && col < ColumnCount)
            _data[row, col] = value;
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
    }

    public void DeleteRow(int row)
    {
        if (row < 0 || row >= RowCount) return;
        for (int r = row; r < RowCount - 1; r++)
            for (int c = 0; c < ColumnCount; c++)
                _data[r, c] = _data[r + 1, c];
        for (int c = 0; c < ColumnCount; c++)
            _data[RowCount - 1, c] = "";
    }

    public void ShiftRowsDown(int fromRow)
    {
        if (fromRow < 0 || fromRow >= RowCount) return;
        // Shift rows down from the bottom, inserting empty row at fromRow
        for (int r = RowCount - 1; r > fromRow; r--)
            for (int c = 0; c < ColumnCount; c++)
                _data[r, c] = _data[r - 1, c];
        for (int c = 0; c < ColumnCount; c++)
            _data[fromRow, c] = "";
    }

    public void ShiftRowsUp(int fromRow)
    {
        if (fromRow < 0 || fromRow >= RowCount) return;
        // Shift rows up, discarding fromRow, empty row at bottom
        for (int r = fromRow; r < RowCount - 1; r++)
            for (int c = 0; c < ColumnCount; c++)
                _data[r, c] = _data[r + 1, c];
        for (int c = 0; c < ColumnCount; c++)
            _data[RowCount - 1, c] = "";
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
        using var writer = new StreamWriter(path);
        for (int r = 0; r < RowCount; r++)
        {
            var fields = new string[ColumnCount];
            for (int c = 0; c < ColumnCount; c++)
                fields[c] = EscapeCsvField(_data[r, c]);
            writer.WriteLine(string.Join(",", fields));
        }
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
}
