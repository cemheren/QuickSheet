namespace ExcelConsole;

public class SpreadsheetApp
{
    private readonly GridManager _grid;
    private int _selectedRow;
    private int _selectedCol;
    private string _clipboard = "";
    private string? _loadedFile;
    private bool _dirty;
    private string? _searchTerm;
    private List<(int row, int col)> _searchMatches = new();
    private int _searchMatchIndex = -1;
    private readonly WebFetchManager _webFetch = new();
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
                string val = _grid.GetCellValue(r, c);
                // Use rendered length for prefixes whose display is shorter than the raw value.
                int len;
                if (CellPrefix.IsSparkline(val))
                    len = CellPrefix.RenderSparkline(val, _grid)?.Length ?? val.Length;
                else if (CellPrefix.IsWebFetch(val))
                {
                    string? url = CellPrefix.ParseWebFetchUrl(val);
                    string display = url != null ? _webFetch.GetDisplay(url) : val;
                    len = display.Length;
                }
                else
                {
                    var cp = CellPrefix.ParseColor(val);
                    len = cp != null ? cp.Value.text.Length : val.Length;
                }
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
            ConsoleKeyInfo key;
            if (_webFetch.HasPending)
            {
                // Polling mode: check for key every 50 ms, re-render to show fetch progress
                if (Console.KeyAvailable)
                {
                    key = Console.ReadKey(intercept: true);
                }
                else
                {
                    Thread.Sleep(50);
                    Render();
                    continue;
                }
            }
            else
            {
                key = Console.ReadKey(intercept: true);
            }

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
                        _dirty = true;
                        if (_selectedRow >= _grid.RowCount) _selectedRow = _grid.RowCount - 1;
                        break;
                    case ConsoleKey.C:
                        _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                        break;
                    case ConsoleKey.X:
                        _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                        _grid.SetCellValue(_selectedRow, _selectedCol, "");
                        _dirty = true;
                        break;
                    case ConsoleKey.V:
                        _grid.SetCellValue(_selectedRow, _selectedCol, _clipboard);
                        _dirty = true;
                        break;
                    case ConsoleKey.O:
                        _grid.ShiftRowsDown(_selectedRow);
                        _dirty = true;
                        break;
                    case ConsoleKey.P:
                        _grid.ShiftRowsUp(_selectedRow);
                        _dirty = true;
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
                    case ConsoleKey.T:
                        Theme.CycleNext();
                        Console.Clear();
                        break;
                    case ConsoleKey.G:
                        PromptGoto();
                        break;
                    case ConsoleKey.Z:
                        _grid.Undo();
                        _dirty = true;
                        break;
                    case ConsoleKey.Y:
                        _grid.Redo();
                        _dirty = true;
                        break;
                    case ConsoleKey.B:
                        _grid.SortByColumn(_selectedCol);
                        _dirty = true;
                        break;
                    case ConsoleKey.R:
                        PromptFindReplace();
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
                    {
                        _grid.SetCellValue(_selectedRow, _selectedCol, val[..^1]);
                        _dirty = true;
                    }
                    break;
                case ConsoleKey.Delete:
                    _grid.SetCellValue(_selectedRow, _selectedCol, "");
                    _dirty = true;
                    break;
                case ConsoleKey.F5:
                    _webFetch.InvalidateAll();
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
                        _dirty = true;
                    }
                    break;
            }

            Render();
        }

        // Auto-save state on exit
        Directory.CreateDirectory(StateDir);
        _grid.SaveToCsv(AutoSavePath);

        _webFetch.Dispose();
        Console.ResetColor();
        Console.CursorVisible = true;
        Console.Clear();
    }

    private void Render()
    {
        var theme = Theme.Current;
        Console.SetCursorPosition(0, 0);
        Console.BackgroundColor = theme.Background;
        Console.ForegroundColor = theme.Foreground;

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
                Console.BackgroundColor = theme.HeaderHighlight;
                Console.Write(header);
                Console.BackgroundColor = theme.Background;
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
                Console.BackgroundColor = theme.HeaderHighlight;
                Console.Write(rowNum);
                Console.BackgroundColor = theme.Background;
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

                // Resolve display text: sparkline > web fetch > color prefix > raw
                string rendered;
                ConsoleColor? colorBg = null;
                if (CellPrefix.IsSparkline(cellVal))
                {
                    rendered = CellPrefix.RenderSparkline(cellVal, _grid) ?? cellVal;
                }
                else if (CellPrefix.IsWebFetch(cellVal))
                {
                    string? url = CellPrefix.ParseWebFetchUrl(cellVal);
                    rendered = url != null ? _webFetch.GetDisplay(url) : cellVal;
                }
                else
                {
                    var colorParsed = CellPrefix.ParseColor(cellVal);
                    if (colorParsed != null)
                    {
                        colorBg = colorParsed.Value.bg;
                        rendered = colorParsed.Value.text;
                    }
                    else
                    {
                        rendered = cellVal;
                    }
                }
                string display = rendered.PadRight(w)[..w];

                bool isSelected = r == _selectedRow && c == _selectedCol;
                bool isMatch = _searchTerm != null && _searchMatches.Contains((r, c));
                if (isSelected && isMatch)
                {
                    Console.BackgroundColor = theme.SearchSelectedBg;
                    Console.ForegroundColor = theme.SearchSelectedFg;
                }
                else if (isSelected)
                {
                    Console.BackgroundColor = theme.SelectionBg;
                    Console.ForegroundColor = theme.SelectionFg;
                }
                else if (isMatch)
                {
                    Console.BackgroundColor = theme.SearchMatchBg;
                    Console.ForegroundColor = theme.SearchMatchFg;
                }
                else if (colorBg != null)
                {
                    Console.BackgroundColor = colorBg.Value;
                    Console.ForegroundColor = ConsoleColor.White;
                }

                Console.Write(display);

                if (isSelected || isMatch || colorBg != null)
                {
                    Console.BackgroundColor = theme.Background;
                    Console.ForegroundColor = theme.Foreground;
                }
            }

            ClearToEndOfLine(gridWidth, totalWidth);
            Console.WriteLine();
        }

        // Status bar at the bottom
        int statusY = Console.WindowHeight - 1;
        Console.SetCursorPosition(0, statusY);
        Console.BackgroundColor = theme.StatusBarBg;
        Console.ForegroundColor = theme.StatusBarFg;

        // File info: name + modified indicator
        string fileLabel = _loadedFile != null ? Path.GetFileName(_loadedFile) : "autosave";
        string dirtyMark = _dirty ? " ●" : "";

        // Count non-empty cells
        int filledCells = 0;
        for (int r = 0; r < _grid.RowCount; r++)
            for (int c = 0; c < _grid.ColumnCount; c++)
                if (!string.IsNullOrEmpty(_grid.GetCellValue(r, c)))
                    filledCells++;

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

        string status = $" {fileLabel}{dirtyMark}  {cellRef}{valueDisplay}{sumDisplay}{productDisplay}{searchDisplay}  │  {filledCells} cells  [{theme.Name}]  Ctrl+H: Help ";
        Console.Write(status.PadRight(totalWidth));

        Console.BackgroundColor = theme.Background;
        Console.ForegroundColor = theme.Foreground;
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
            "  ║  Ctrl+R         Find & Replace           ║",
            "  ║  Ctrl+G         Go to cell (e.g. A1)     ║",
            "  ║  Ctrl+Z         Undo                     ║",
            "  ║  Ctrl+Y         Redo                     ║",
            "  ║  Ctrl+B         Sort by column (toggle)  ║",
            "  ║  Enter          Next search match         ║",
            "  ║  Shift+Enter    Previous search match     ║",
            "  ║  Escape         Clear search              ║",
            "  ║  Ctrl+T         Cycle theme               ║",
            "  ║  Ctrl+H         Show this help            ║",
            "  ║  Ctrl+Q         Quit                     ║",
            "  ║                                          ║",
            "  ║  F5             Refresh w: web cells     ║",
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

    private void PromptFindReplace()
    {
        int statusY = Console.WindowHeight - 1;
        int totalWidth = Console.WindowWidth;

        // Step 1: prompt for search term
        string? findTerm = PromptInput(" Find: ", statusY, totalWidth);
        if (string.IsNullOrEmpty(findTerm)) return;

        // Count matches
        int matchCount = 0;
        for (int r = 0; r < _grid.RowCount; r++)
            for (int c = 0; c < _grid.ColumnCount; c++)
                if (_grid.GetCellValue(r, c).Contains(findTerm, StringComparison.OrdinalIgnoreCase))
                    matchCount++;

        if (matchCount == 0)
        {
            Console.SetCursorPosition(0, statusY);
            Console.BackgroundColor = ConsoleColor.DarkRed;
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write($" No matches for \"{findTerm}\"".PadRight(totalWidth));
            Console.ResetColor();
            Console.ReadKey(intercept: true);
            return;
        }

        // Step 2: prompt for replacement
        string? replaceTerm = PromptInput($" Replace ({matchCount} matches) with: ", statusY, totalWidth);
        if (replaceTerm == null) return; // Escape pressed

        // Step 3: confirm
        Console.SetCursorPosition(0, statusY);
        Console.BackgroundColor = ConsoleColor.Yellow;
        Console.ForegroundColor = ConsoleColor.Black;
        Console.Write($" Replace {matchCount} occurrence(s)? [y/n] ".PadRight(totalWidth));
        Console.ResetColor();

        var confirm = Console.ReadKey(intercept: true);
        if (confirm.KeyChar != 'y' && confirm.KeyChar != 'Y') return;

        // Step 4: perform replacement
        int replaced = 0;
        for (int r = 0; r < _grid.RowCount; r++)
        {
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                string val = _grid.GetCellValue(r, c);
                if (val.Contains(findTerm, StringComparison.OrdinalIgnoreCase))
                {
                    string newVal = val.Replace(findTerm, replaceTerm, StringComparison.OrdinalIgnoreCase);
                    _grid.SetCellValue(r, c, newVal);
                    replaced++;
                }
            }
        }

        _dirty = true;

        Console.SetCursorPosition(0, statusY);
        Console.BackgroundColor = ConsoleColor.DarkGreen;
        Console.ForegroundColor = ConsoleColor.White;
        Console.Write($" Replaced {replaced} cell(s)".PadRight(totalWidth));
        Console.ResetColor();
        Console.ReadKey(intercept: true);
    }

    private string? PromptInput(string prompt, int statusY, int totalWidth)
    {
        Console.SetCursorPosition(0, statusY);
        Console.BackgroundColor = ConsoleColor.White;
        Console.ForegroundColor = ConsoleColor.Black;
        Console.Write(prompt.PadRight(totalWidth));
        Console.SetCursorPosition(prompt.Length, statusY);
        Console.CursorVisible = true;

        string input = "";
        while (true)
        {
            var k = Console.ReadKey(intercept: true);
            if (k.Key == ConsoleKey.Enter) break;
            if (k.Key == ConsoleKey.Escape) { Console.CursorVisible = false; Console.ResetColor(); return null; }
            if (k.Key == ConsoleKey.Backspace)
            {
                if (input.Length > 0)
                {
                    input = input[..^1];
                    Console.SetCursorPosition(prompt.Length, statusY);
                    Console.Write(input.PadRight(totalWidth - prompt.Length));
                    Console.SetCursorPosition(prompt.Length + input.Length, statusY);
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
        Console.ResetColor();
        return input;
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
        _dirty = false;

        // Flash confirmation
        Console.SetCursorPosition(0, statusY);
        Console.Write($" Saved to {filename} ".PadRight(totalWidth));
        Console.BackgroundColor = ConsoleColor.Black;
        Console.ForegroundColor = ConsoleColor.White;
        Thread.Sleep(1000);
    }

    private void PromptGoto()
    {
        int statusY = Console.WindowHeight - 1;
        Console.SetCursorPosition(0, statusY);
        Console.BackgroundColor = ConsoleColor.White;
        Console.ForegroundColor = ConsoleColor.Black;
        int totalWidth = Console.WindowWidth;

        Console.Write(" Go to cell (e.g. A1, C5): ".PadRight(totalWidth));
        Console.SetCursorPosition(" Go to cell (e.g. A1, C5): ".Length, statusY);
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
                    Console.SetCursorPosition(" Go to cell (e.g. A1, C5): ".Length, statusY);
                    Console.Write(input.PadRight(totalWidth - " Go to cell (e.g. A1, C5): ".Length));
                    Console.SetCursorPosition(" Go to cell (e.g. A1, C5): ".Length + input.Length, statusY);
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

        var parsed = CellPrefix.ParseCellRef(input);
        if (parsed is var (row, col))
        {
            _selectedRow = Math.Clamp(row, 0, _grid.RowCount - 1);
            _selectedCol = Math.Clamp(col, 0, _grid.ColumnCount - 1);
        }
        else
        {
            // Flash error
            Console.SetCursorPosition(0, statusY);
            Console.BackgroundColor = ConsoleColor.Red;
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write($" Invalid cell reference: {input} ".PadRight(totalWidth));
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.White;
            Thread.Sleep(1000);
        }
    }
}
