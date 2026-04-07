using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Borderless form that replaces the desktop experience. Covers the working area
/// (behind the taskbar) and resists minimization so Win+D reveals it instead of
/// the normal desktop. Fully interactable — click to focus, type to edit cells.
/// </summary>
internal class DesktopForm : Form
{
    private readonly GridManager _grid;
    private int _selectedRow;
    private int _selectedCol;
    private readonly HashSet<(int row, int col)> _selection = new();
    private string _clipboard = "";
    private string? _loadedFile;

    private bool _editing;
    private string _editText = "";
    private int _editCursorPos;

    private readonly Font _monoFont;
    private readonly int _charWidth;
    private readonly int _charHeight;

    private readonly NotifyIcon _trayIcon;
    private bool _lockZOrder;
    private IntPtr _winEventHook;
    private NativeMethods.WinEventDelegate? _winEventDelegate;
    private System.Threading.Timer? _autoSaveTimer;

    private readonly int _colWidth;
    private const int RowHeaderWidth = 4;

    private static readonly string StateDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelConsole");
    private static readonly string AutoSavePath = Path.Combine(StateDir, "autosave.csv");

    public DesktopForm(string? csvPath)
    {
        Text = "QuickSheet";
        BackColor = Color.Black;
        FormBorderStyle = FormBorderStyle.None;
        Opacity = 0.85;
        StartPosition = FormStartPosition.Manual;
        ShowInTaskbar = false;
        KeyPreview = true;
        SetStyle(
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.UserPaint, true);

        // Fill the working area (screen minus taskbar)
        Bounds = Screen.PrimaryScreen!.WorkingArea;

        _monoFont = new Font("Consolas", 14f, FontStyle.Regular);
        using (var bmp = new Bitmap(1, 1))
        using (var g = Graphics.FromImage(bmp))
        {
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            var size = TextRenderer.MeasureText(g, "MMMMMMMMMM", _monoFont,
                new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding);
            _charWidth = (int)Math.Ceiling(size.Width / 10.0);
            _charHeight = size.Height;
        }

        int availableWidth = Bounds.Width / _charWidth;
        int availableHeight = Bounds.Height / _charHeight - 3;
        _grid = new GridManager(availableWidth, availableHeight);

        // Distribute available width evenly so columns fill the screen
        int usableChars = availableWidth - RowHeaderWidth;
        _colWidth = _grid.ColumnCount > 0 ? usableChars / _grid.ColumnCount : 20;

        if (csvPath is not null)
        {
            _loadedFile = csvPath;
            _grid.LoadFromCsv(csvPath);
        }
        else if (File.Exists(AutoSavePath))
        {
            _grid.LoadFromCsv(AutoSavePath);
        }

        PopulateDesktopFiles();

        _trayIcon = new NotifyIcon
        {
            Text = "QuickSheet Desktop",
            Icon = SystemIcons.Application,
            Visible = true,
        };
        var menu = new ContextMenuStrip();
        menu.Items.Add("Save", null, (_, _) => SaveFile());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) => Close());
        _trayIcon.ContextMenuStrip = menu;

        KeyDown += OnFormKeyDown;
        KeyPress += OnFormKeyPress;
        MouseDown += OnFormMouseDown;
        MouseDoubleClick += OnFormDoubleClick;
        FormClosing += OnFormClosing;

        // Autosave every 5 seconds in case of crash/termination
        Directory.CreateDirectory(StateDir);
        _autoSaveTimer = new System.Threading.Timer(_ =>
        {
            try { _grid.SaveToCsv(AutoSavePath); } catch { }
        }, null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(5));
    }

    // ── Desktop files ───────────────────────────────────────────────

    private void PopulateDesktopFiles()
    {
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (!Directory.Exists(desktopPath)) return;

        var entries = Directory.GetFileSystemEntries(desktopPath)
            .OrderBy(e => Directory.Exists(e) ? 0 : 1)  // folders first
            .ThenBy(Path.GetFileName)
            .ToArray();

        // Place files in the last column, flowing to previous columns if needed
        int col = _grid.ColumnCount - 1;
        int row = 0;
        foreach (var entry in entries)
        {
            if (row >= _grid.RowCount)
            {
                row = 0;
                col--;
                if (col < 0) break;
            }
            string name = Path.GetFileName(entry);
            if (Directory.Exists(entry)) name = "\ud83d\udcc1 " + name;
            _grid.SetFileEntry(row, col, name, entry);
            row++;
        }
    }

    private void OpenSelectedFile()
    {
        if (!_grid.IsFileEntry(_selectedRow, _selectedCol)) return;
        string path = _grid.GetFilePath(_selectedRow, _selectedCol);
        if (string.IsNullOrEmpty(path)) return;
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch { }
    }

    private static bool IsHyperlink(string value) =>
        value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        value.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

    private static bool IsCommand(string value) =>
        value.StartsWith("r: ", StringComparison.OrdinalIgnoreCase);

    private static (string exe, string args) ParseCommand(string value)
    {
        // Strip the "r: " prefix
        string cmd = value[3..].Trim();
        // Support quoted executable: r: cmd "qs autosave.csv" or unquoted: r: notepad
        if (cmd.Length == 0) return ("", "");
        if (cmd[0] == '"')
        {
            int close = cmd.IndexOf('"', 1);
            if (close > 0)
                return (cmd[1..close], cmd[(close + 1)..].Trim());
        }
        int space = cmd.IndexOf(' ');
        if (space < 0) return (cmd, "");
        return (cmd[..space], cmd[(space + 1)..].Trim());
    }

    private static void RunCommand(string cellValue)
    {
        if (!IsCommand(cellValue)) return;
        var (exe, args) = ParseCommand(cellValue);
        if (string.IsNullOrEmpty(exe)) return;
        try
        {
            Process.Start(new ProcessStartInfo(exe, args)
            {
                UseShellExecute = true
            });
        }
        catch { }
    }

    private void OpenHyperlink()
    {
        string url = _grid.GetCellValue(_selectedRow, _selectedCol);
        if (!IsHyperlink(url)) return;
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch { }
    }

    private void OpenAllSelected()
    {
        // Collect all cells: current cursor + multi-selection
        var cells = new HashSet<(int row, int col)>(_selection)
        {
            (_selectedRow, _selectedCol)
        };

        bool opened = false;
        foreach (var (r, c) in cells)
        {
            if (_grid.IsFileEntry(r, c))
            {
                string path = _grid.GetFilePath(r, c);
                if (!string.IsNullOrEmpty(path))
                    try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); opened = true; } catch { }
            }
            else
            {
                string val = _grid.GetCellValue(r, c);
                if (IsCommand(val))
                    { RunCommand(val); opened = true; }
                else if (IsHyperlink(val))
                    try { Process.Start(new ProcessStartInfo(val) { UseShellExecute = true }); opened = true; } catch { }
            }
        }

        // If nothing was opened, move down as default Enter behavior
        if (!opened && _selectedRow < _grid.RowCount - 1)
            _selectedRow++;
    }

    // ── Desktop mode ─────────────────────────────────────────────────

    /// <summary>
    /// Configures the form as a desktop replacement: hidden from Alt+Tab.
    /// </summary>
    public void EnterDesktopMode()
    {
        // Hide from Alt+Tab
        var exStyle = (long)NativeMethods.GetWindowLongPtr(Handle, NativeMethods.GWL_EXSTYLE);
        exStyle |= NativeMethods.WS_EX_TOOLWINDOW;
        NativeMethods.SetWindowLongPtr(Handle, NativeMethods.GWL_EXSTYLE, (IntPtr)exStyle);

        // Send to bottom of Z-order so it starts behind all existing windows
        NativeMethods.SetWindowPos(Handle, NativeMethods.HWND_BOTTOM,
            0, 0, 0, 0,
            NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);

        // Lock Z-order so clicking doesn't bring it above other apps
        _lockZOrder = true;

        // Hook foreground window changes to detect Win+D ("Show Desktop").
        // When WorkerW becomes the foreground, Windows has issued "Show Desktop"
        // — make ourselves topmost so we appear above it. When any other window
        // becomes foreground, drop back to non-topmost.
        _winEventDelegate = OnForegroundChanged;
        _winEventHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_SYSTEM_FOREGROUND,
            NativeMethods.EVENT_SYSTEM_FOREGROUND,
            IntPtr.Zero, _winEventDelegate, 0, 0,
            NativeMethods.WINEVENT_OUTOFCONTEXT);

        Invalidate();
    }

    private void OnForegroundChanged(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        var sb = new System.Text.StringBuilder(32);
        NativeMethods.GetClassName(hwnd, sb, sb.Capacity);
        string cls = sb.ToString();

        // Temporarily unlock Z-order so SetWindowPos can change it
        _lockZOrder = false;

        if (cls is "WorkerW" or "Progman")
        {
            // "Show Desktop" was triggered — become topmost to appear above it
            NativeMethods.SetWindowPos(Handle, NativeMethods.HWND_TOPMOST,
                0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
        }
        else
        {
            // A real app took focus — drop topmost so we go behind apps again
            NativeMethods.SetWindowPos(Handle, NativeMethods.HWND_NOTOPMOST,
                0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
        }

        _lockZOrder = true;
    }

    protected override bool ShowWithoutActivation => true;

    protected override void WndProc(ref Message m)
    {
        // Prevent the form from rising in Z-order (covers other apps) while
        // still allowing activation for keyboard focus.
        if (m.Msg == NativeMethods.WM_WINDOWPOSCHANGING && _lockZOrder)
        {
            var pos = Marshal.PtrToStructure<NativeMethods.WINDOWPOS>(m.LParam);
            pos.flags |= NativeMethods.SWP_NOZORDER;
            Marshal.StructureToPtr(pos, m.LParam, false);
        }
        base.WndProc(ref m);
    }

    // ── Rendering ────────────────────────────────────────────────────

    private int[] GetColumnWidths()
    {
        var widths = new int[_grid.ColumnCount];
        for (int c = 0; c < _grid.ColumnCount; c++)
            widths[c] = _colWidth;
        return widths;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        RenderGrid(g, ClientSize.Width, ClientSize.Height);
    }

    private void RenderGrid(Graphics g, int formWidth, int formHeight)
    {
        int cw = _charWidth;
        int ch = _charHeight;
        int[] colWidths = GetColumnWidths();
        g.Clear(Color.Black);
        int y = 0;

        // Column headers
        DrawText(g, new string(' ', RowHeaderWidth), 0, y, Color.White, Color.Black);
        int x = RowHeaderWidth * cw;
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            int w = colWidths[c];
            string header = GridManager.GetColumnName(c).PadRight(w);
            Color bg = c == _selectedCol ? Color.FromArgb(64, 64, 64) : Color.Black;
            DrawText(g, header, x, y, Color.White, bg);
            x += w * cw;
        }
        y += ch;

        // Underline
        string underline = new string('-', RowHeaderWidth);
        for (int c = 0; c < _grid.ColumnCount; c++)
            underline += new string('-', colWidths[c]);
        DrawText(g, underline, 0, y, Color.White, Color.Black);
        y += ch;

        // Data rows
        int statusY = formHeight - ch;
        for (int r = 0; r < _grid.RowCount && y + ch <= statusY; r++)
        {
            x = 0;
            string rowNum = (r + 1).ToString().PadLeft(RowHeaderWidth - 1) + " ";
            Color rowBg = r == _selectedRow ? Color.FromArgb(64, 64, 64) : Color.Black;
            DrawText(g, rowNum, x, y, Color.White, rowBg);
            x = RowHeaderWidth * cw;
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                int w = colWidths[c];
                string cellVal = _grid.GetCellValue(r, c);
                string display = cellVal.Length >= w ? cellVal[..w] : cellVal.PadRight(w);
                bool isCursor = r == _selectedRow && c == _selectedCol;
                bool isMultiSel = _selection.Contains((r, c));
                bool isFile = _grid.IsFileEntry(r, c);
                bool isLink = IsHyperlink(cellVal);
                bool isCmd = IsCommand(cellVal);
                Color bg = isCursor   ? Color.FromArgb(64, 64, 64)
                         : isMultiSel ? Color.FromArgb(50, 50, 80)
                         : isFile     ? Color.FromArgb(0, 40, 60)
                         : isLink     ? Color.FromArgb(40, 0, 60)
                         : isCmd      ? Color.FromArgb(40, 40, 0)
                         : Color.Black;
                Color fg = isFile ? Color.FromArgb(100, 200, 255)
                         : isLink ? Color.FromArgb(180, 140, 255)
                         : isCmd  ? Color.FromArgb(255, 220, 100)
                         : Color.White;
                DrawText(g, display, x, y, fg, bg);
                x += w * cw;
            }
            y += ch;
        }

        // Status bar
        int maxChars = formWidth / cw;
        string status;
        if (_editing)
        {
            string prefix = " Edit: ";
            string cursor = _editText.Insert(_editCursorPos, "\u2502");
            status = $"{prefix}{cursor}  (Enter=OK  Esc=Cancel)";
        }
        else
        {
            string cellRef = _grid.GetCellReference(_selectedRow, _selectedCol);
            string value = _grid.GetCellValue(_selectedRow, _selectedCol);
            string valueDisplay = string.IsNullOrEmpty(value) ? "" : $" = {value}";
            double? sum = _grid.GetColumnSum(_selectedCol);
            string colName = GridManager.GetColumnName(_selectedCol);
            string sumDisplay = sum.HasValue ? $"  \u03a3{colName} = {sum.Value}" : "";
            double? product = _grid.GetRowProduct(_selectedRow);
            string productDisplay = product.HasValue ? $"  \u03a0{(_selectedRow + 1)} = {product.Value}" : "";
            status = $" {cellRef}{valueDisplay}{sumDisplay}{productDisplay}  |  F2: Edit  Ctrl+S: Save  Ctrl+Q: Quit";
        }
        status = status.PadRight(maxChars);
        g.FillRectangle(Brushes.White, 0, statusY, formWidth, ch);
        DrawText(g, status, 0, statusY, Color.Black, Color.White);
    }

    private void DrawText(Graphics g, string text, int x, int y, Color fg, Color bg)
    {
        int w = text.Length * _charWidth;
        using var bgBrush = new SolidBrush(bg);
        g.FillRectangle(bgBrush, x, y, w, _charHeight);
        TextRenderer.DrawText(g, text, _monoFont, new Point(x, y), fg, bg,
            TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
    }

    // ── Input ────────────────────────────────────────────────────────

    private void EnterEditMode()
    {
        _editing = true;
        _editText = _grid.GetCellValue(_selectedRow, _selectedCol);
        _editCursorPos = _editText.Length;
    }

    private void CommitEdit()
    {
        _grid.SetCellValue(_selectedRow, _selectedCol, _editText);
        _editing = false;
    }

    private void CancelEdit()
    {
        _editing = false;
    }

    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        if (_editing)
        {
            bool handled = true;
            switch (e.KeyCode)
            {
                case Keys.Left:
                    if (_editCursorPos > 0) _editCursorPos--;
                    break;
                case Keys.Right:
                    if (_editCursorPos < _editText.Length) _editCursorPos++;
                    break;
                case Keys.Home:
                    _editCursorPos = 0;
                    break;
                case Keys.End:
                    _editCursorPos = _editText.Length;
                    break;
                case Keys.Back:
                    if (_editCursorPos > 0)
                    {
                        _editText = _editText.Remove(_editCursorPos - 1, 1);
                        _editCursorPos--;
                    }
                    break;
                case Keys.Delete:
                    if (_editCursorPos < _editText.Length)
                        _editText = _editText.Remove(_editCursorPos, 1);
                    break;
                case Keys.Enter:
                    CommitEdit();
                    break;
                case Keys.Escape:
                    CancelEdit();
                    break;
                default:
                    handled = false;
                    break;
            }
            if (handled)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                Invalidate();
            }
            return;
        }

        bool handled2 = true;
        if (e.Control)
        {
            switch (e.KeyCode)
            {
                case Keys.Q: Close(); return;
                case Keys.D:
                    _grid.DeleteRow(_selectedRow);
                    if (_selectedRow >= _grid.RowCount) _selectedRow = _grid.RowCount - 1;
                    break;
                case Keys.C:
                    if (_selection.Count > 0)
                    {
                        var copyCells = new List<(int row, int col)>(_selection) { (_selectedRow, _selectedCol) };
                        copyCells.Sort((a, b) => a.row != b.row ? a.row.CompareTo(b.row) : a.col.CompareTo(b.col));
                        _clipboard = string.Join(Environment.NewLine, copyCells.Select(c => _grid.GetCellValue(c.row, c.col)));
                    }
                    else
                        _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                    if (!string.IsNullOrEmpty(_clipboard))
                        Clipboard.SetText(_clipboard);
                    break;
                case Keys.X:
                    if (_selection.Count > 0)
                    {
                        var cutCells = new List<(int row, int col)>(_selection) { (_selectedRow, _selectedCol) };
                        cutCells.Sort((a, b) => a.row != b.row ? a.row.CompareTo(b.row) : a.col.CompareTo(b.col));
                        _clipboard = string.Join(Environment.NewLine, cutCells.Select(c => _grid.GetCellValue(c.row, c.col)));
                        foreach (var (r, c) in cutCells) _grid.SetCellValue(r, c, "");
                        _selection.Clear();
                    }
                    else
                    {
                        _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                        _grid.SetCellValue(_selectedRow, _selectedCol, "");
                    }
                    if (!string.IsNullOrEmpty(_clipboard))
                        Clipboard.SetText(_clipboard);
                    break;
                case Keys.V:
                    string paste = Clipboard.ContainsText() ? Clipboard.GetText() : _clipboard;
                    string[] lines = paste.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        int targetRow = _selectedRow + i;
                        if (targetRow >= _grid.RowCount) break;
                        _grid.SetCellValue(targetRow, _selectedCol, lines[i]);
                    }
                    break;
                case Keys.O: _grid.ShiftRowsDown(_selectedRow); break;
                case Keys.P: _grid.ShiftRowsUp(_selectedRow); break;
                case Keys.S: SaveFile(); break;
                default: handled2 = false; break;
            }
        }
        else if (e.Shift)
        {
            // Shift+Arrow: extend multi-selection
            _selection.Add((_selectedRow, _selectedCol));
            switch (e.KeyCode)
            {
                case Keys.Up: if (_selectedRow > 0) _selectedRow--; break;
                case Keys.Down: if (_selectedRow < _grid.RowCount - 1) _selectedRow++; break;
                case Keys.Left: if (_selectedCol > 0) _selectedCol--; break;
                case Keys.Right: if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++; break;
                default: handled2 = false; break;
            }
            if (handled2) _selection.Add((_selectedRow, _selectedCol));
        }
        else
        {
            switch (e.KeyCode)
            {
                case Keys.Up:
                    if (_selectedRow > 0) _selectedRow--;
                    _selection.Clear();
                    break;
                case Keys.Down:
                    if (_selectedRow < _grid.RowCount - 1) _selectedRow++;
                    _selection.Clear();
                    break;
                case Keys.Left:
                    if (_selectedCol > 0) _selectedCol--;
                    _selection.Clear();
                    break;
                case Keys.Right:
                    if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                    _selection.Clear();
                    break;
                case Keys.Enter:
                    OpenAllSelected();
                    break;
                case Keys.F2:
                    EnterEditMode();
                    break;
                case Keys.Back:
                    var val = _grid.GetCellValue(_selectedRow, _selectedCol);
                    if (val.Length > 0) _grid.SetCellValue(_selectedRow, _selectedCol, val[..^1]);
                    break;
                case Keys.Delete: _grid.SetCellValue(_selectedRow, _selectedCol, ""); break;
                case Keys.Tab:
                    if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                    else if (_selectedRow < _grid.RowCount - 1) { _selectedCol = 0; _selectedRow++; }
                    _selection.Clear();
                    break;
                case Keys.Escape:
                    _selection.Clear();
                    break;
                default: handled2 = false; break;
            }
        }
        if (handled2)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            Invalidate();
        }
    }

    private void OnFormKeyPress(object? sender, KeyPressEventArgs e)
    {
        if (e.KeyChar >= 32 && e.KeyChar <= 126)
        {
            if (_editing)
            {
                _editText = _editText.Insert(_editCursorPos, e.KeyChar.ToString());
                _editCursorPos++;
            }
            else
            {
                var cur = _grid.GetCellValue(_selectedRow, _selectedCol);
                _grid.SetCellValue(_selectedRow, _selectedCol, cur + e.KeyChar);
            }
            e.Handled = true;
            Invalidate();
        }
    }

    private void OnFormMouseDown(object? sender, MouseEventArgs e)
    {
        int headerRows = 2; // column header + underline
        int row = (e.Y / _charHeight) - headerRows;
        if (row < 0 || row >= _grid.RowCount) return;

        int[] colWidths = GetColumnWidths();
        int x = RowHeaderWidth * _charWidth;
        int col = -1;
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            int colPx = colWidths[c] * _charWidth;
            if (e.X >= x && e.X < x + colPx) { col = c; break; }
            x += colPx;
        }
        if (col < 0) return;

        if (ModifierKeys.HasFlag(Keys.Control))
        {
            // Ctrl+click: toggle cell in multi-selection
            if (!_selection.Remove((row, col)))
                _selection.Add((row, col));
        }
        else
        {
            _selection.Clear();
        }

        _selectedRow = row;
        _selectedCol = col;
        Invalidate();
    }

    private void OnFormDoubleClick(object? sender, MouseEventArgs e)
    {
        OpenAllSelected();
    }

    // ── File operations ──────────────────────────────────────────────

    private void SaveFile()
    {
        string filename = _loadedFile ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "spreadsheet.csv");
        _grid.SaveToCsv(filename);
        _loadedFile = filename;
        _trayIcon.ShowBalloonTip(2000, "QuickSheet", $"Saved to {filename}", ToolTipIcon.Info);
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_winEventHook != IntPtr.Zero)
            NativeMethods.UnhookWinEvent(_winEventHook);
        Directory.CreateDirectory(StateDir);
        _grid.SaveToCsv(AutoSavePath);
        _trayIcon.Visible = false;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _autoSaveTimer?.Dispose();
            _monoFont.Dispose();
            _trayIcon.Dispose();
        }
        base.Dispose(disposing);
    }
}
