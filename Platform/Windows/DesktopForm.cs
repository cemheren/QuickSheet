using System.Drawing;
using System.Drawing.Text;
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
    private string _clipboard = "";
    private string? _loadedFile;

    private readonly Font _monoFont;
    private readonly int _charWidth;
    private readonly int _charHeight;

    private readonly NotifyIcon _trayIcon;
    private bool _lockZOrder;

    private const int DefaultColWidth = 20;
    private const int RowHeaderWidth = 4;

    private static readonly string StateDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelConsole");
    private static readonly string AutoSavePath = Path.Combine(StateDir, "autosave.csv");

    public DesktopForm(string? csvPath)
    {
        Text = "QuickSheet";
        BackColor = Color.Black;
        FormBorderStyle = FormBorderStyle.None;
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

        if (csvPath is not null)
        {
            _loadedFile = csvPath;
            _grid.LoadFromCsv(csvPath);
        }
        else if (File.Exists(AutoSavePath))
        {
            _grid.LoadFromCsv(AutoSavePath);
        }

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
        FormClosing += OnFormClosing;
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

        // Now lock Z-order so it stays there
        _lockZOrder = true;
        Invalidate();
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
        {
            int max = GridManager.GetColumnName(c).Length;
            for (int r = 0; r < _grid.RowCount; r++)
            {
                int len = _grid.GetCellValue(r, c).Length;
                if (len > max) max = len;
            }
            widths[c] = Math.Max(DefaultColWidth, max + 2);
        }
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
                bool isSelected = r == _selectedRow && c == _selectedCol;
                Color bg = isSelected ? Color.FromArgb(64, 64, 64) : Color.Black;
                DrawText(g, display, x, y, Color.White, bg);
                x += w * cw;
            }
            y += ch;
        }

        // Status bar
        string cellRef = _grid.GetCellReference(_selectedRow, _selectedCol);
        string value = _grid.GetCellValue(_selectedRow, _selectedCol);
        string valueDisplay = string.IsNullOrEmpty(value) ? "" : $" = {value}";
        double? sum = _grid.GetColumnSum(_selectedCol);
        string colName = GridManager.GetColumnName(_selectedCol);
        string sumDisplay = sum.HasValue ? $"  \u03a3{colName} = {sum.Value}" : "";
        double? product = _grid.GetRowProduct(_selectedRow);
        string productDisplay = product.HasValue ? $"  \u03a0{(_selectedRow + 1)} = {product.Value}" : "";
        string status = $" {cellRef}{valueDisplay}{sumDisplay}{productDisplay}  |  Ctrl+S: Save  Ctrl+Q: Quit";
        int maxChars = formWidth / cw;
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

    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        bool handled = true;
        if (e.Control)
        {
            switch (e.KeyCode)
            {
                case Keys.Q: Close(); return;
                case Keys.D:
                    _grid.DeleteRow(_selectedRow);
                    if (_selectedRow >= _grid.RowCount) _selectedRow = _grid.RowCount - 1;
                    break;
                case Keys.C: _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol); break;
                case Keys.X:
                    _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                    _grid.SetCellValue(_selectedRow, _selectedCol, "");
                    break;
                case Keys.V: _grid.SetCellValue(_selectedRow, _selectedCol, _clipboard); break;
                case Keys.O: _grid.ShiftRowsDown(_selectedRow); break;
                case Keys.P: _grid.ShiftRowsUp(_selectedRow); break;
                case Keys.S: SaveFile(); break;
                default: handled = false; break;
            }
        }
        else
        {
            switch (e.KeyCode)
            {
                case Keys.Up: if (_selectedRow > 0) _selectedRow--; break;
                case Keys.Down: if (_selectedRow < _grid.RowCount - 1) _selectedRow++; break;
                case Keys.Left: if (_selectedCol > 0) _selectedCol--; break;
                case Keys.Right: if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++; break;
                case Keys.Enter: if (_selectedRow < _grid.RowCount - 1) _selectedRow++; break;
                case Keys.Back:
                    var val = _grid.GetCellValue(_selectedRow, _selectedCol);
                    if (val.Length > 0) _grid.SetCellValue(_selectedRow, _selectedCol, val[..^1]);
                    break;
                case Keys.Delete: _grid.SetCellValue(_selectedRow, _selectedCol, ""); break;
                case Keys.Tab:
                    if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                    else if (_selectedRow < _grid.RowCount - 1) { _selectedCol = 0; _selectedRow++; }
                    break;
                case Keys.Escape: break;
                default: handled = false; break;
            }
        }
        if (handled)
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
            var cur = _grid.GetCellValue(_selectedRow, _selectedCol);
            _grid.SetCellValue(_selectedRow, _selectedCol, cur + e.KeyChar);
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

        _selectedRow = row;
        _selectedCol = col;
        Invalidate();
    }

    // ── File operations ──────────────────────────────────────────────

    private void SaveFile()
    {
        string filename = _loadedFile ?? "spreadsheet.csv";
        _grid.SaveToCsv(filename);
        _loadedFile = filename;
        _trayIcon.ShowBalloonTip(2000, "QuickSheet", $"Saved to {filename}", ToolTipIcon.Info);
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        Directory.CreateDirectory(StateDir);
        _grid.SaveToCsv(AutoSavePath);
        _trayIcon.Visible = false;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _monoFont.Dispose();
            _trayIcon.Dispose();
        }
        base.Dispose(disposing);
    }
}
