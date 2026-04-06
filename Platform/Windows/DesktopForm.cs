using System.Drawing;
using System.Drawing.Text;
using System.Windows.Forms;

namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Borderless fullscreen form that renders the spreadsheet grid using GDI+.
/// Designed to be embedded into the Windows desktop WorkerW layer as a live wallpaper,
/// or popped out for interactive editing via Alt+` hotkey.
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

    private bool _isEmbedded;
    private IntPtr _workerW;
    private readonly NotifyIcon _trayIcon;

    private const int MinColWidth = 10;
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
        Bounds = Screen.PrimaryScreen!.Bounds;
        ShowInTaskbar = false;
        KeyPreview = true;
        SetStyle(
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.UserPaint, true);

        // Monospace font — measure character cell size
        _monoFont = new Font("Consolas", 12f, FontStyle.Regular);
        using (var bmp = new Bitmap(1, 1))
        using (var g = Graphics.FromImage(bmp))
        {
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            var size = TextRenderer.MeasureText(g, "MMMMMMMMMM", _monoFont,
                new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding);
            _charWidth = (int)Math.Ceiling(size.Width / 10.0);
            _charHeight = size.Height;
        }

        // Create grid sized to fill the screen
        int availableWidth = Bounds.Width / _charWidth;
        int availableHeight = Bounds.Height / _charHeight - 3; // header, underline, status bar
        _grid = new GridManager(availableWidth, availableHeight);

        // Load data
        if (csvPath is not null)
        {
            _loadedFile = csvPath;
            _grid.LoadFromCsv(csvPath);
        }
        else if (File.Exists(AutoSavePath))
        {
            _grid.LoadFromCsv(AutoSavePath);
        }

        // System tray icon
        _trayIcon = new NotifyIcon
        {
            Text = "QuickSheet Desktop",
            Icon = SystemIcons.Application,
            Visible = true,
        };
        var menu = new ContextMenuStrip();
        menu.Items.Add("Toggle Focus (Alt+`)", null, (_, _) => ToggleEmbed());
        menu.Items.Add("Save", null, (_, _) => SaveFile());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) => Close());
        _trayIcon.ContextMenuStrip = menu;
        _trayIcon.DoubleClick += (_, _) => ToggleEmbed();

        KeyDown += OnFormKeyDown;
        KeyPress += OnFormKeyPress;
        Paint += OnFormPaint;
        Resize += (_, _) => Invalidate();
        FormClosing += OnFormClosing;
    }

    // ── Desktop embedding ────────────────────────────────────────────

    public void EmbedIntoDesktop()
    {
        _workerW = DesktopEmbedder.GetDesktopWorkerW();
        if (_workerW == IntPtr.Zero)
        {
            // Fallback: run as a normal borderless window
            _isEmbedded = false;
            return;
        }

        NativeMethods.SetParent(Handle, _workerW);
        // As a child of WorkerW, position is relative to parent — origin at (0,0)
        Location = Point.Empty;
        Size = Screen.PrimaryScreen!.Bounds.Size;
        _isEmbedded = true;

        NativeMethods.RegisterHotKey(Handle, NativeMethods.HOTKEY_ID,
            NativeMethods.MOD_ALT | NativeMethods.MOD_NOREPEAT, NativeMethods.VK_OEM_3);
    }

    private void ToggleEmbed()
    {
        if (_isEmbedded)
        {
            // Pop out for interactive editing — detach from WorkerW, use screen coordinates
            NativeMethods.SetParent(Handle, IntPtr.Zero);
            Location = Screen.PrimaryScreen!.Bounds.Location;
            Size = Screen.PrimaryScreen!.Bounds.Size;
            TopMost = true;
            NativeMethods.SetForegroundWindow(Handle);
            _isEmbedded = false;
            Focus();
        }
        else if (_workerW != IntPtr.Zero)
        {
            // Sink back behind desktop icons — child coordinates relative to WorkerW
            TopMost = false;
            NativeMethods.SetParent(Handle, _workerW);
            Location = Point.Empty;
            Size = Screen.PrimaryScreen!.Bounds.Size;
            _isEmbedded = true;
        }
        Invalidate();
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == NativeMethods.WM_HOTKEY && m.WParam.ToInt32() == NativeMethods.HOTKEY_ID)
        {
            ToggleEmbed();
            return;
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
            widths[c] = Math.Max(MinColWidth, max + 2);
        }
        return widths;
    }

    private void OnFormPaint(object? sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        int cw = _charWidth;
        int ch = _charHeight;
        int[] colWidths = GetColumnWidths();

        int y = 0;

        // ── Column headers ──
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
        ClearToRight(g, x, y, ch);
        y += ch;

        // ── Underline ──
        string underline = new string('─', RowHeaderWidth);
        for (int c = 0; c < _grid.ColumnCount; c++)
            underline += new string('─', colWidths[c]);
        DrawText(g, underline, 0, y, Color.White, Color.Black);
        ClearToRight(g, underline.Length * cw, y, ch);
        y += ch;

        // ── Data rows ──
        int statusY = Height - ch;
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

            ClearToRight(g, x, y, ch);
            y += ch;
        }

        // Clear empty area between last row and status bar
        if (y < statusY)
            g.FillRectangle(Brushes.Black, 0, y, Width, statusY - y);

        // ── Status bar ──
        string cellRef = _grid.GetCellReference(_selectedRow, _selectedCol);
        string value = _grid.GetCellValue(_selectedRow, _selectedCol);
        string valueDisplay = string.IsNullOrEmpty(value) ? "" : $" = {value}";

        double? sum = _grid.GetColumnSum(_selectedCol);
        string colName = GridManager.GetColumnName(_selectedCol);
        string sumDisplay = sum.HasValue ? $"  \u03a3{colName} = {sum.Value}" : "";

        double? product = _grid.GetRowProduct(_selectedRow);
        string productDisplay = product.HasValue ? $"  \u03a0{_selectedRow + 1} = {product.Value}" : "";

        string mode = _isEmbedded ? "DESKTOP" : "FOCUSED";
        string status = $" {cellRef}{valueDisplay}{sumDisplay}{productDisplay}  \u2502  Alt+`: Toggle  Ctrl+Q: Quit  [{mode}]";
        int maxChars = Width / cw;
        status = status.PadRight(maxChars);

        g.FillRectangle(Brushes.White, 0, statusY, Width, ch);
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

    private void ClearToRight(Graphics g, int fromX, int y, int rowHeight)
    {
        if (fromX < Width)
            g.FillRectangle(Brushes.Black, fromX, y, Width - fromX, rowHeight);
    }

    // ── Input ────────────────────────────────────────────────────────

    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        bool handled = true;

        if (e.Control)
        {
            switch (e.KeyCode)
            {
                case Keys.Q:
                    Close();
                    return;
                case Keys.D:
                    _grid.DeleteRow(_selectedRow);
                    if (_selectedRow >= _grid.RowCount) _selectedRow = _grid.RowCount - 1;
                    break;
                case Keys.C:
                    _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                    break;
                case Keys.X:
                    _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                    _grid.SetCellValue(_selectedRow, _selectedCol, "");
                    break;
                case Keys.V:
                    _grid.SetCellValue(_selectedRow, _selectedCol, _clipboard);
                    break;
                case Keys.O:
                    _grid.ShiftRowsDown(_selectedRow);
                    break;
                case Keys.P:
                    _grid.ShiftRowsUp(_selectedRow);
                    break;
                case Keys.S:
                    SaveFile();
                    break;
                default:
                    handled = false;
                    break;
            }
        }
        else
        {
            switch (e.KeyCode)
            {
                case Keys.Up:
                    if (_selectedRow > 0) _selectedRow--;
                    break;
                case Keys.Down:
                    if (_selectedRow < _grid.RowCount - 1) _selectedRow++;
                    break;
                case Keys.Left:
                    if (_selectedCol > 0) _selectedCol--;
                    break;
                case Keys.Right:
                    if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                    break;
                case Keys.Enter:
                    if (_selectedRow < _grid.RowCount - 1) _selectedRow++;
                    break;
                case Keys.Back:
                    var val = _grid.GetCellValue(_selectedRow, _selectedCol);
                    if (val.Length > 0)
                        _grid.SetCellValue(_selectedRow, _selectedCol, val[..^1]);
                    break;
                case Keys.Delete:
                    _grid.SetCellValue(_selectedRow, _selectedCol, "");
                    break;
                case Keys.Tab:
                    if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                    else if (_selectedRow < _grid.RowCount - 1) { _selectedCol = 0; _selectedRow++; }
                    break;
                case Keys.Escape:
                    break;
                default:
                    handled = false;
                    break;
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
        NativeMethods.UnregisterHotKey(Handle, NativeMethods.HOTKEY_ID);
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
