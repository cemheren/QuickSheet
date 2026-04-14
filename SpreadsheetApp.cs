namespace ExcelConsole;

public class SpreadsheetApp
{
    private readonly GridManager _grid;
    private int _selectedRow;
    private int _selectedCol;
    private string _clipboard = "";
    private string? _loadedFile;
    private string? _searchTerm;
    private List<(int row, int col)> _searchMatches = new();
    private int _searchMatchIndex = -1;
    private const int MinColWidth = 10;
    private const int RowHeaderWidth = 4;

    private static readonly string StateDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelConsole");
    private static readonly string AutoSavePath = Path.Combine(StateDir, "autosave.csv");

    public SpreadsheetApp(string? csvPath = null)
    {
        int w = Console.WindowWidth;
        int h = Console.WindowHeight;
        _grid = new GridManager(w, h - 3);
        if (csvPath is not null)
        {
            _loadedFile = csvPath;
            _grid.LoadFromCsv(csvPath);
        }
        else if (File.Exists(AutoSavePath))
        {
            _grid.LoadFromCsv(AutoSavePath);
        }
    }

    private int[] GetColumnWidths()
    {
        var widths = new int[_grid.ColumnCount];
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            int max = GridManager.GetColumnName(c).Length;
            for (int r = 0; r < _grid.RowCount; r++)
            {
                int len = _grid.GetCellValue(r, c).Length;
                if (len > max) max = len;
            }
            widths[c] = Math.Max(MinColWidth, max + 2); // +2 for padding
        }
        return widths;
    }

    public void Run()
    {
        Console.CursorVisible = false;
        Console.TreatControlCAsInput = true;
        Console.Clear();
        Render();

        while (true)
        {
            var key = Console.ReadKey(intercept: true);

            if (key.Modifiers.HasFlag(ConsoleModifiers.Control) && key.Key == ConsoleKey.Q)
                break;

            // While search results active, only allow navigation/clear
            if (_searchTerm != null)
            {
                if (key.Key == ConsoleKey.Enter)
                {
                    if (key.Modifiers.HasFlag(ConsoleModifiers.Shift))
                        _searchMatchIndex = (_searchMatchIndex - 1 + _searchMatches.Count) % _searchMatches.Count;
                    else
                        _searchMatchIndex = (_searchMatchIndex + 1) % _searchMatches.Count;
                    if (_searchMatches.Count > 0)
                    {
                        _selectedRow = _searchMatches[_searchMatchIndex].row;
                        _selectedCol = _searchMatches[_searchMatchIndex].col;
                    }
                }
                else if (key.Key == ConsoleKey.Escape)
                {
                    _searchTerm = null;
                    _searchMatches.Clear();
                    _searchMatchIndex = -1;
                }
                Render();
                continue;
            }

            // Ctrl shortcuts
            if (key.Modifiers.HasFlag(ConsoleModifiers.Control))
            {
                switch (key.Key)
                {
                    case ConsoleKey.D:
                        _grid.DeleteRow(_selectedRow);
                        if (_selectedRow >= _grid.RowCount) _selectedRow = _grid.RowCount - 1;
                        break;
                    case ConsoleKey.C:
                        _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                        break;
                    case ConsoleKey.X:
                        _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                        _grid.SetCellValue(_selectedRow, _selectedCol, "");
                        break;
                    case ConsoleKey.V:
                        _grid.SetCellValue(_selectedRow, _selectedCol, _clipboard);
                        break;
                    case ConsoleKey.O:
                        _grid.ShiftRowsDown(_selectedRow);
                        break;
                    case ConsoleKey.P:
                        _grid.ShiftRowsUp(_selectedRow);
                        break;
                    case ConsoleKey.H:
                    case ConsoleKey.Backspace:
                        ShowHelp();
                        break;
                    case ConsoleKey.S:
                        PromptSave();
                        break;
                    case ConsoleKey.F:
                        PromptSearch();
                        break;
                }
                Render();
                continue;
            }

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    if (_selectedRow > 0) _selectedRow--;
                    break;
                case ConsoleKey.DownArrow:
                    if (_selectedRow < _grid.RowCount - 1) _selectedRow++;
                    break;
                case ConsoleKey.LeftArrow:
                    if (_selectedCol > 0) _selectedCol--;
                    break;
                case ConsoleKey.RightArrow:
                    if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                    break;
                case ConsoleKey.Enter:
                    if (_selectedRow < _grid.RowCount - 1) _selectedRow++;
                    break;
                case ConsoleKey.Backspace:
                    var val = _grid.GetCellValue(_selectedRow, _selectedCol);
                    if (val.Length > 0)
                        _grid.SetCellValue(_selectedRow, _selectedCol, val[..^1]);
                    break;
                case ConsoleKey.Delete:
                    _grid.SetCellValue(_selectedRow, _selectedCol, "");
                    break;
                case ConsoleKey.Escape:
                    break;
                case ConsoleKey.Tab:
                    if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                    else if (_selectedRow < _grid.RowCount - 1) { _selectedCol = 0; _selectedRow++; }
                    break;
                default:
                    if (key.KeyChar >= 32 && key.KeyChar <= 126)
                    {
                        var cur = _grid.GetCellValue(_selectedRow, _selectedCol);
                        _grid.SetCellValue(_selectedRow, _selectedCol, cur + key.KeyChar);
                    }
                    break;
            }

            Render();
        }

        // Auto-save state on exit
        Directory.CreateDirectory(StateDir);
        _grid.SaveToCsv(AutoSavePath);

        Console.ResetColor();
        Console.CursorVisible = true;
        Console.Clear();
    }

    private void Render()
    {
        Console.SetCursorPosition(0, 0);
        Console.BackgroundColor = ConsoleColor.Black;
        Console.ForegroundColor = ConsoleColor.White;

        int totalWidth = Console.WindowWidth;
        int[] colWidths = GetColumnWidths();
        int gridWidth = RowHeaderWidth;
        for (int c = 0; c < _grid.ColumnCount; c++) gridWidth += colWidths[c];

        // Header row
        Console.Write(new string(' ', RowHeaderWidth));
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            int w = colWidths[c];
            string header = GridManager.GetColumnName(c).PadRight(w);
            if (c == _selectedCol)
            {
                Console.BackgroundColor = ConsoleColor.DarkGray;
                Console.Write(header);
                Console.BackgroundColor = ConsoleColor.Black;
            }
            else
            {
                Console.Write(header);
            }
        }
        ClearToEndOfLine(gridWidth, totalWidth);
        Console.WriteLine();

        // Header underline
        Console.Write(new string('─', RowHeaderWidth));
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            Console.Write(new string('─', colWidths[c]));
        }
        ClearToEndOfLine(gridWidth, totalWidth);
        Console.WriteLine();

        // Data rows
        for (int r = 0; r < _grid.RowCount; r++)
        {
            // Row number
            string rowNum = (r + 1).ToString().PadLeft(RowHeaderWidth - 1) + " ";
            if (r == _selectedRow)
            {
                Console.BackgroundColor = ConsoleColor.DarkGray;
                Console.Write(rowNum);
                Console.BackgroundColor = ConsoleColor.Black;
            }
            else
            {
                Console.Write(rowNum);
            }

            // Cell values
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                int w = colWidths[c];
                string cellVal = _grid.GetCellValue(r, c);
                string display = cellVal.PadRight(w)[..w];

                bool isSelected = r == _selectedRow && c == _selectedCol;
                bool isMatch = _searchTerm != null && _searchMatches.Contains((r, c));
                if (isSelected && isMatch)
                {
                    Console.BackgroundColor = ConsoleColor.Green;
                    Console.ForegroundColor = ConsoleColor.Black;
                }
                else if (isSelected)
                {
                    Console.BackgroundColor = ConsoleColor.DarkGray;
                    Console.ForegroundColor = ConsoleColor.White;
                }
                else if (isMatch)
                {
                    Console.BackgroundColor = ConsoleColor.DarkYellow;
                    Console.ForegroundColor = ConsoleColor.Black;
                }

                Console.Write(display);

                if (isSelected || isMatch)
                {
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }

            ClearToEndOfLine(gridWidth, totalWidth);
            Console.WriteLine();
        }

        // Status bar at the bottom
        int statusY = Console.WindowHeight - 1;
        Console.SetCursorPosition(0, statusY);
        Console.BackgroundColor = ConsoleColor.White;
        Console.ForegroundColor = ConsoleColor.Black;

        string cellRef = _grid.GetCellReference(_selectedRow, _selectedCol);
        string value = _grid.GetCellValue(_selectedRow, _selectedCol);
        string valueDisplay = string.IsNullOrEmpty(value) ? "" : $" = {value}";

        double? sum = _grid.GetColumnSum(_selectedCol);
        string colName = GridManager.GetColumnName(_selectedCol);
        string sumDisplay = sum.HasValue ? $"  Σ{colName} = {sum.Value}" : "";

        double? product = _grid.GetRowProduct(_selectedRow);
        string productDisplay = product.HasValue ? $"  Π{_selectedRow + 1} = {product.Value}" : "";

        string searchDisplay = _searchTerm != null
            ? $"  🔍\"{_searchTerm}\" {(_searchMatches.Count > 0 ? $"{_searchMatchIndex + 1}/{_searchMatches.Count}" : "no matches")}"
            : "";

        string status = $" {cellRef}{valueDisplay}{sumDisplay}{productDisplay}{searchDisplay}  │  Ctrl+Q: Quit ";
        Console.Write(status.PadRight(totalWidth));

        Console.BackgroundColor = ConsoleColor.Black;
        Console.ForegroundColor = ConsoleColor.White;
    }

    private static void ClearToEndOfLine(int currentPos, int totalWidth)
    {
        int remaining = totalWidth - currentPos;
        if (remaining > 0)
            Console.Write(new string(' ', remaining));
    }

    private static void ShowHelp()
    {
        Console.Clear();
        Console.BackgroundColor = ConsoleColor.Black;
        Console.ForegroundColor = ConsoleColor.White;
        Console.SetCursorPosition(0, 0);

        string[] lines =
        [
            "",
            "  ╔══════════════════════════════════════════╗",
            "  ║         ExcelConsole — Shortcuts         ║",
            "  ╠══════════════════════════════════════════╣",
            "  ║                                          ║",
            "  ║  Arrow Keys     Navigate cells           ║",
            "  ║  Tab            Next cell                ║",
            "  ║  Enter          Move down one row        ║",
            "  ║  Backspace      Delete last character    ║",
            "  ║  Delete         Clear cell               ║",
            "  ║                                          ║",
            "  ║  Ctrl+C         Copy cell                ║",
            "  ║  Ctrl+X         Cut cell                 ║",
            "  ║  Ctrl+V         Paste cell               ║",
            "  ║  Ctrl+D         Delete row               ║",
            "  ║  Ctrl+O         Insert row (shift down)  ║",
            "  ║  Ctrl+P         Remove row (shift up)    ║",
            "  ║  Ctrl+S         Save to CSV              ║",
            "  ║  Ctrl+F         Find (contains search)   ║",
            "  ║  Enter          Next search match         ║",
            "  ║  Shift+Enter    Previous search match     ║",
            "  ║  Escape         Clear search              ║",
            "  ║  Ctrl+H         Show this help            ║",
            "  ║  Ctrl+Q         Quit                     ║",
            "  ║                                          ║",
            "  ║  Usage: dotnet run [file.csv]            ║",
            "  ║                                          ║",
            "  ╚══════════════════════════════════════════╝",
            "",
            "  Press any key to return...",
        ];

        foreach (var line in lines)
            Console.WriteLine(line);

        Console.ReadKey(intercept: true);
    }

    private void PromptSearch()
    {
        int statusY = Console.WindowHeight - 1;
        Console.SetCursorPosition(0, statusY);
        Console.BackgroundColor = ConsoleColor.White;
        Console.ForegroundColor = ConsoleColor.Black;
        int totalWidth = Console.WindowWidth;

        Console.Write(" Find: ".PadRight(totalWidth));
        Console.SetCursorPosition(" Find: ".Length, statusY);
        Console.CursorVisible = true;

        string input = "";
        while (true)
        {
            var k = Console.ReadKey(intercept: true);
            if (k.Key == ConsoleKey.Enter) break;
            if (k.Key == ConsoleKey.Escape)
            {
                Console.CursorVisible = false;
                return;
            }
            if (k.Key == ConsoleKey.Backspace)
            {
                if (input.Length > 0)
                {
                    input = input[..^1];
                    Console.SetCursorPosition(" Find: ".Length, statusY);
                    Console.Write(input.PadRight(totalWidth - " Find: ".Length));
                    Console.SetCursorPosition(" Find: ".Length + input.Length, statusY);
                }
                continue;
            }
            if (k.KeyChar >= 32 && k.KeyChar <= 126)
            {
                input += k.KeyChar;
                Console.Write(k.KeyChar);
            }
        }

        Console.CursorVisible = false;

        if (string.IsNullOrEmpty(input))
        {
            _searchTerm = null;
            _searchMatches.Clear();
            _searchMatchIndex = -1;
            return;
        }

        _searchTerm = input;
        _searchMatches.Clear();
        for (int r = 0; r < _grid.RowCount; r++)
            for (int c = 0; c < _grid.ColumnCount; c++)
                if (_grid.GetCellValue(r, c).Contains(input, StringComparison.OrdinalIgnoreCase))
                    _searchMatches.Add((r, c));

        if (_searchMatches.Count > 0)
        {
            _searchMatchIndex = 0;
            _selectedRow = _searchMatches[0].row;
            _selectedCol = _searchMatches[0].col;
        }
        else
        {
            _searchMatchIndex = -1;
        }
    }

    private void PromptSave()
    {
        int statusY = Console.WindowHeight - 1;
        Console.SetCursorPosition(0, statusY);
        Console.BackgroundColor = ConsoleColor.White;
        Console.ForegroundColor = ConsoleColor.Black;
        int totalWidth = Console.WindowWidth;

        string defaultName = _loadedFile ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "spreadsheet.csv");
        Console.Write($" Save as [{defaultName}]: ".PadRight(totalWidth));
        Console.SetCursorPosition($" Save as [{defaultName}]: ".Length, statusY);
        Console.CursorVisible = true;

        string input = "";
        while (true)
        {
            var k = Console.ReadKey(intercept: true);
            if (k.Key == ConsoleKey.Enter) break;
            if (k.Key == ConsoleKey.Escape) { Console.CursorVisible = false; return; }
            if (k.Key == ConsoleKey.Backspace)
            {
                if (input.Length > 0)
                {
                    input = input[..^1];
                    Console.SetCursorPosition($" Save as [{defaultName}]: ".Length, statusY);
                    Console.Write(input.PadRight(totalWidth - $" Save as [{defaultName}]: ".Length));
                    Console.SetCursorPosition($" Save as [{defaultName}]: ".Length + input.Length, statusY);
                }
                continue;
            }
            if (k.KeyChar >= 32 && k.KeyChar <= 126)
            {
                input += k.KeyChar;
                Console.Write(k.KeyChar);
            }
        }

        Console.CursorVisible = false;
        string filename = string.IsNullOrWhiteSpace(input) ? defaultName : input;
        if (!filename.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            filename += ".csv";

        _grid.SaveToCsv(filename);
        _loadedFile = filename;

        // Flash confirmation
        Console.SetCursorPosition(0, statusY);
        Console.Write($" Saved to {filename} ".PadRight(totalWidth));
        Console.BackgroundColor = ConsoleColor.Black;
        Console.ForegroundColor = ConsoleColor.White;
        Thread.Sleep(1000);
    }
}
