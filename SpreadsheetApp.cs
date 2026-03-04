namespace ExcelConsole;

public class SpreadsheetApp
{
    private readonly GridManager _grid;
    private int _selectedRow;
    private int _selectedCol;
    private string _clipboard = "";
    private const int MinColWidth = 10;
    private const int RowHeaderWidth = 4;

    public SpreadsheetApp()
    {
        int w = Console.WindowWidth;
        int h = Console.WindowHeight;
        // Reserve: 1 header row, 1 header underline, 1 status bar
        _grid = new GridManager(w, h - 3);
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
                    case ConsoleKey.Backspace: // Ctrl+H sends Backspace on Linux
                        ShowHelp();
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
                if (isSelected)
                {
                    Console.BackgroundColor = ConsoleColor.DarkGray;
                    Console.ForegroundColor = ConsoleColor.White;
                }

                Console.Write(display);

                if (isSelected)
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

        string status = $" {cellRef}{valueDisplay}{sumDisplay}{productDisplay}  │  Ctrl+Q: Quit ";
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
            "  ║  Ctrl+H         Show this help           ║",
            "  ║  Ctrl+Q         Quit                     ║",
            "  ║                                          ║",
            "  ╚══════════════════════════════════════════╝",
            "",
            "  Press any key to return...",
        ];

        foreach (var line in lines)
            Console.WriteLine(line);

        Console.ReadKey(intercept: true);
    }
}
