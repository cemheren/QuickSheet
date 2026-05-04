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
internal class DesktopForm : DesktopFormBase
{
    private GridManager _grid;

    //todo: refactor these two into the GridManager
    private readonly HashSet<(int row, int col)> _selection = new();
    private string? _loadedFile;

    //todo: why is this even here. 
    private string _clipboard = "";

    //todo: refactor into a SearchingMode : IMode. 
    private bool _searching;
    private string _searchInput = "";
    private string? _searchTerm;
    private List<(int row, int col)> _searchMatches = new();
    private int _searchMatchIndex = -1;

    private readonly EditingMode _editMode;

    //todo: DraggingMode? Not sure if this should be a state yet. 
    private bool _dragging;
    private int _dragAnchorRow;
    private int _dragAnchorCol;

    private Font _monoFont;
    private int _charWidth;
    private int _charHeight;

    private readonly NotifyIcon _trayIcon;
    private bool _showResolved;
    private System.Threading.Timer? _autoSaveTimer;
    private System.Threading.Timer? _csvReloadTimer;

    private int _colWidth;
    private int _columnWidth = 20;
    private const int MinColumnWidth = 8;
    private const int MaxColumnWidth = 40;
    private const int ColumnWidthStep = 4;
    private const int RowHeaderWidth = 4;

    private readonly InlineProcessManager _processManager = new();
    private readonly Extensions.ExtensionManager _extensionManager;
    private System.Threading.Timer? _inlineRefreshTimer;
    private readonly LoopManager _loopManager;

    private static readonly string AutoSavePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "autosave.csv");

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

        UpdateFontMetrics();

        int availableWidth = Bounds.Width / _charWidth;
        int availableHeight = Bounds.Height / _charHeight - 3;
        _grid = new GridManager(availableWidth, availableHeight, _columnWidth);
        _editMode = new EditingMode(_grid);
        _extensionManager = new Extensions.ExtensionManager(new WindowsExtensionEnvironment(), _grid);
        _loopManager = new LoopManager(_grid, ActivateCellAt);

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
        MouseMove += OnFormMouseMove;
        MouseUp += OnFormMouseUp;
        MouseDoubleClick += OnFormDoubleClick;
        FormClosing += OnFormClosing;

        // Autosave every 5 seconds if there are pending changes
        _autoSaveTimer = new System.Threading.Timer(_ =>
        {
            try { if (_grid.IsDirty) _grid.SaveToCsv(AutoSavePath); } catch { }
        }, null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(5));

        // Reload CSV every 60 seconds to pick up external (OneDrive) changes
        _csvReloadTimer = new System.Threading.Timer(_ =>
        {
            try
            {
                string? path = _loadedFile ?? (File.Exists(AutoSavePath) ? AutoSavePath : null);
                if (path is not null && _grid.MergeFromCsv(path))
                    BeginInvoke(() => Invalidate());
            }
            catch { }
        }, null, TimeSpan.FromSeconds(60), TimeSpan.FromSeconds(60));

        // Refresh inline process output every 500ms
        _inlineRefreshTimer = new System.Threading.Timer(_ =>
        {
            try
            {
                _extensionManager.ScanGrid();
                if (_processManager.HasAnyNewOutput() || _extensionManager.ConsumeHasChanges())
                    BeginInvoke(() => Invalidate());
            }
            catch { }
        }, null, TimeSpan.FromMilliseconds(500), TimeSpan.FromMilliseconds(500));
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

    private void UpdateFontMetrics()
    {
        float fontSize = Math.Clamp(_columnWidth * 0.7f, 6f, 22f);
        _monoFont?.Dispose();
        _monoFont = new Font("Consolas", fontSize, FontStyle.Regular);
        using var bmp = new Bitmap(1, 1);
        using var g = Graphics.FromImage(bmp);
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        var size = TextRenderer.MeasureText(g, "MMMMMMMMMM", _monoFont,
            new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding);
        _charWidth = (int)Math.Ceiling(size.Width / 10.0);
        _charHeight = size.Height;
    }

    private void RebuildGrid()
    {
        // Save only if there are unsaved user edits
        try { if (_grid.IsDirty) _grid.SaveToCsv(_loadedFile ?? AutoSavePath); } catch { }

        UpdateFontMetrics();
        Bounds = Screen.PrimaryScreen!.WorkingArea;
        int availableWidth = Bounds.Width / _charWidth;
        int availableHeight = Bounds.Height / _charHeight - 3;
        var newGrid = new GridManager(availableWidth, availableHeight, _columnWidth);

        int usableChars = availableWidth - RowHeaderWidth;
        _colWidth = newGrid.ColumnCount > 0 ? usableChars / newGrid.ColumnCount : 20;

        // Copy user data (skip file entries — they'll be re-populated at new positions)
        for (int r = 0; r < Math.Min(_grid.RowCount, newGrid.RowCount); r++)
            for (int c = 0; c < Math.Min(_grid.ColumnCount, newGrid.ColumnCount); c++)
                if (!_grid.IsFileEntry(r, c))
                    newGrid.SetCellValue(r, c, _grid.GetCellValue(r, c));

        var (prevRow, prevCol) = _grid.GetCurrentCell();
        _grid = newGrid;
        _editMode.Grid = _grid;
        _extensionManager.UpdateGrid(_grid);
        _loopManager.UpdateGrid(_grid);
        _grid.SelectCell(prevRow, prevCol);

        if (_loadedFile is not null)
            _grid.LoadFromCsv(_loadedFile);
        else if (File.Exists(AutoSavePath))
            _grid.LoadFromCsv(AutoSavePath);

        PopulateDesktopFiles();

        _selection.Clear();
        Invalidate();
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

    /// <summary>
    /// If the cell is an inline ref pointing to a command, kill the cached process
    /// so it re-runs on the next render. Returns true if it handled the action.
    /// </summary>
    private bool TryRerunInlineCommand(int row, int col)
    {
        string val = _grid.GetCellValue(row, col);
        if (!CellPrefix.IsInline(val)) return false;

        string? resolved = _grid.ResolveInline(row, col);
        if (resolved == null || !CellPrefix.IsCommand(resolved)) return false;

        _processManager.StopProcess(row, col);
        Invalidate();
        return true;
    }

    private void OpenAllSelected()
    {
        // Collect all cells: current cursor + multi-selection
        var (curRow, curCol) = _grid.GetCurrentCell();
        var cells = new HashSet<(int row, int col)>(_selection)
        {
            (curRow, curCol)
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
                { 
                    RunCommand(val); 
                    opened = true; 
                }
                else if (IsHyperlink(val))
                    try { Process.Start(new ProcessStartInfo(val) { UseShellExecute = true }); opened = true; } catch { }
            }
        }

        // If nothing was opened, move down as default Enter behavior
        if (!opened)
            _grid.MoveDown();
    }

    /// <summary>
    /// Activates a single cell as if the user pressed Enter on it.
    /// Used by LoopManager for periodic activation of target cells.
    /// </summary>
    private void ActivateCellAt(int row, int col)
    {
        if (row < 0 || row >= _grid.RowCount || col < 0 || col >= _grid.ColumnCount)
            return;

        _extensionManager.ReactivateCell(row, col);

        string val = _grid.GetCellValue(row, col);
        if (CellPrefix.IsInline(val))
        {
            string? resolved = _grid.ResolveInline(row, col);
            if (resolved != null && CellPrefix.IsCommand(resolved))
            {
                _processManager.StopProcess(row, col);
                try { BeginInvoke(() => Invalidate()); } catch { }
                return;
            }
        }

        if (_grid.IsFileEntry(row, col))
        {
            string path = _grid.GetFilePath(row, col);
            if (!string.IsNullOrEmpty(path))
                try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); } catch { }
        }
        else if (IsCommand(val))
        {
            RunCommand(val);
        }
        else if (IsHyperlink(val))
        {
            try { Process.Start(new ProcessStartInfo(val) { UseShellExecute = true }); } catch { }
        }
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
        var (selRow, selCol) = _grid.GetCurrentCell();
        int cw = _charWidth;
        int ch = _charHeight;
        int[] colWidths = GetColumnWidths();
        g.Clear(Color.Black);
        int y = 0;

        // Track inline cells that need process management
        var activeInlineCmds = new HashSet<(int, int)>();

        // Pre-scan for inline cells with visual spans (i: A10,5,3)
        var inlineSpans = new List<(int anchorRow, int anchorCol, int spanCols, int spanRows, string content, bool isCmd)>();
        var occludedCells = new HashSet<(int, int)>();
        for (int r = 0; r < _grid.RowCount; r++)
        {
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                string val = _grid.GetCellValue(r, c);
                if (!CellPrefix.IsInline(val)) continue;
                string expandedVal = CellPrefix.ExpandCellReferences(val, _grid);
                var parsed = CellPrefix.ParseInlineRef(expandedVal);
                if (parsed == null || (parsed.Value.spanCols <= 1 && parsed.Value.spanRows <= 1)) continue;

                int sc = parsed.Value.spanCols, sr = parsed.Value.spanRows;
                int endCol = Math.Min(c + sc - 1, _grid.ColumnCount - 1);
                int endRow = Math.Min(r + sr - 1, _grid.RowCount - 1);

                // Resolve content
                string? resolved = _grid.ResolveInline(r, c);
                bool isCmd = false;
                string content;
                if (resolved != null && CellPrefix.IsCommand(resolved))
                {
                    isCmd = true;
                    activeInlineCmds.Add((r, c));
                    string expandedCmd = CellPrefix.ExpandCellReferences(resolved, _grid);
                    // Calculate pty size from span dimensions (subtract padding)
                    int ptyCols = -1; // account for 8px padding → ~1 char
                    for (int tc = c; tc <= endCol; tc++)
                        ptyCols += colWidths[tc];
                    if (ptyCols < 20) ptyCols = 20;
                    int ptyRows = Math.Max(1, (endRow - r + 1) - 1); // minus header row
                    _processManager.EnsureRunning(r, c, expandedCmd, ptyCols, ptyRows);
                    content = _processManager.GetOutput(r, c) ?? "[running...]";
                }
                else if (resolved != null)
                {
                    content = CellPrefix.ExpandCellReferences(resolved, _grid);
                }
                else
                {
                    content = val;
                }

                inlineSpans.Add((r, c, sc, sr, content, isCmd));

                // Mark all spanned cells except anchor as occluded
                for (int wr = r; wr <= endRow; wr++)
                    for (int wc = c; wc <= endCol; wc++)
                        if (wr != r || wc != c)
                            occludedCells.Add((wr, wc));
            }
        }

        // Column headers
        DrawText(g, new string(' ', RowHeaderWidth), 0, y, Color.White, Color.Black);
        int x = RowHeaderWidth * cw;
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            int w = colWidths[c];
            string header = GridManager.GetColumnName(c).PadRight(w);
            Color bg = c == selCol ? Color.FromArgb(64, 64, 64) : Color.Black;
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
            Color rowBg = r == selRow ? Color.FromArgb(64, 64, 64) : Color.Black;
            DrawText(g, rowNum, x, y, Color.White, rowBg);
            x = RowHeaderWidth * cw;
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                int w = colWidths[c];

                // Skip cells occluded by an inline span
                if (occludedCells.Contains((r, c)))
                {
                    x += w * cw;
                    continue;
                }

                string cellVal = _grid.GetCellValue(r, c);
                string displayVal = cellVal;
                bool isInline = CellPrefix.IsInline(cellVal);
                bool isInlineCmd = false;

                // Resolve inline references
                if (isInline)
                {
                    string? resolved = _grid.ResolveInline(r, c);
                    if (resolved != null)
                    {
                        // If resolved value is a command, show process output
                        if (CellPrefix.IsCommand(resolved))
                        {
                            isInlineCmd = true;
                            activeInlineCmds.Add((r, c));
                            string expandedCmd = CellPrefix.ExpandCellReferences(resolved, _grid);
                            _processManager.EnsureRunning(r, c, expandedCmd);
                            displayVal = _processManager.GetOutput(r, c) ?? "[running...]";
                        }
                        else
                        {
                            displayVal = CellPrefix.ExpandCellReferences(resolved, _grid);
                        }
                    }
                }
                else
                {
                    // Expand cell references in normal cells
                    displayVal = CellPrefix.ExpandCellReferences(cellVal, _grid);
                }

                string display = displayVal.Length >= w ? displayVal[..w] : displayVal.PadRight(w);
                bool isCursor = r == selRow && c == selCol;
                bool isMultiSel = _selection.Contains((r, c));
                bool isSearchMatch = _searchTerm != null && _searchMatches.Contains((r, c));
                bool isFile = _grid.IsFileEntry(r, c);
                bool isLink = IsHyperlink(cellVal);
                bool isCmd = IsCommand(cellVal);
                bool isConflict = cellVal.StartsWith("c: ", StringComparison.Ordinal);
                var extStatus = _extensionManager.GetCellStatus(r, c);
                Color bg = isCursor && isSearchMatch ? Color.FromArgb(0, 180, 0)
                         : isCursor   ? Color.FromArgb(64, 64, 64)
                         : isMultiSel ? Color.FromArgb(50, 50, 80)
                         : isSearchMatch ? Color.FromArgb(80, 80, 0)
                         : isConflict ? Color.FromArgb(100, 0, 0)
                         : isInlineCmd ? Color.FromArgb(20, 50, 20)
                         : isInline   ? Color.FromArgb(0, 40, 50)
                         : isFile     ? Color.FromArgb(0, 40, 60)
                         : isLink     ? Color.FromArgb(40, 0, 60)
                         : isCmd      ? Color.FromArgb(40, 40, 0)
                         : extStatus == Extensions.ExtensionCellStatus.Error ? Color.FromArgb(50, 10, 10)
                         : extStatus == Extensions.ExtensionCellStatus.Running ? Color.FromArgb(10, 40, 10)
                         : Color.Black;
                Color fg = isConflict ? Color.FromArgb(255, 180, 180)
                         : isInlineCmd ? Color.FromArgb(100, 255, 150)
                         : isInline   ? Color.FromArgb(100, 220, 240)
                         : isFile ? Color.FromArgb(100, 200, 255)
                         : isLink ? Color.FromArgb(180, 140, 255)
                         : isCmd  ? Color.FromArgb(255, 220, 100)
                         : extStatus == Extensions.ExtensionCellStatus.Error ? Color.FromArgb(255, 80, 80)
                         : extStatus == Extensions.ExtensionCellStatus.Running ? Color.FromArgb(80, 255, 80)
                         : Color.White;
                DrawText(g, display, x, y, fg, bg);
                x += w * cw;
            }
            y += ch;
        }

        // Draw inline span overlays
        int headerRows = 2;
        foreach (var span in inlineSpans)
        {
            int spanX = RowHeaderWidth * cw;
            for (int c = 0; c < span.anchorCol && c < _grid.ColumnCount; c++)
                spanX += colWidths[c] * cw;

            int spanY = headerRows * ch + span.anchorRow * ch;
            int spanWidth = 0;
            int endCol = Math.Min(span.anchorCol + span.spanCols - 1, _grid.ColumnCount - 1);
            for (int c = span.anchorCol; c <= endCol; c++)
                spanWidth += colWidths[c] * cw;
            int endRow = Math.Min(span.anchorRow + span.spanRows - 1, _grid.RowCount - 1);
            int spanHeight = (endRow - span.anchorRow + 1) * ch;

            if (spanY + spanHeight > statusY) spanHeight = statusY - spanY;
            if (spanWidth <= 0 || spanHeight <= 0) continue;

            // Background
            bool hasCursor = selRow >= span.anchorRow && selRow <= endRow &&
                             selCol >= span.anchorCol && selCol <= endCol;
            Color spanBg = span.isCmd
                ? (hasCursor ? Color.FromArgb(30, 60, 30) : Color.FromArgb(20, 50, 20))
                : (hasCursor ? Color.FromArgb(10, 55, 65) : Color.FromArgb(0, 40, 50));
            using (var bgBrush = new SolidBrush(spanBg))
                g.FillRectangle(bgBrush, spanX, spanY, spanWidth, spanHeight);

            // Border
            Color borderColor = hasCursor ? Color.FromArgb(100, 180, 200) : Color.FromArgb(60, 100, 120);
            using (var borderPen = new Pen(borderColor, 1))
                g.DrawRectangle(borderPen, spanX, spanY, spanWidth - 1, spanHeight - 1);

            // Cell ref label
            string cellRef = _grid.GetCellReference(span.anchorRow, span.anchorCol);
            DrawText(g, cellRef, spanX + 2, spanY + 1, Color.FromArgb(80, 140, 160), spanBg);

            // Render content — strip markdown, show last lines that fit
            if (!string.IsNullOrEmpty(span.content))
            {
                int contentTop = spanY + ch + 2;
                int contentHeight = spanHeight - ch - 4;
                int contentWidth = spanWidth - 8;
                if (contentHeight > 0 && contentWidth > 0)
                {
                    string cleaned = StripMarkdown(span.content);
                    string[] allLines = cleaned.Split('\n');

                    Color contentFg = span.isCmd ? Color.FromArgb(100, 255, 150) : Color.FromArgb(180, 220, 255);

                    // Wrap long lines to fit width, then take last N
                    int charsPerLine = Math.Max(1, contentWidth / cw);
                    var wrappedLines = new List<string>();
                    foreach (string rawLine in allLines)
                    {
                        if (rawLine.Length <= charsPerLine)
                        {
                            wrappedLines.Add(rawLine);
                        }
                        else
                        {
                            // Word-wrap at charsPerLine
                            for (int pos = 0; pos < rawLine.Length; pos += charsPerLine)
                                wrappedLines.Add(rawLine.Substring(pos, Math.Min(charsPerLine, rawLine.Length - pos)));
                        }
                    }

                    int linesPerWindow = Math.Max(1, contentHeight / ch);
                    int startIdx = Math.Max(0, wrappedLines.Count - linesPerWindow);
                    int count = Math.Min(linesPerWindow, wrappedLines.Count);

                    int lineY = contentTop;
                    for (int li = startIdx; li < startIdx + count && lineY + ch <= contentTop + contentHeight; li++)
                    {
                        string line = wrappedLines[li];
                        if (line.Length > charsPerLine)
                            line = line[..charsPerLine];
                        DrawText(g, line.PadRight(charsPerLine), spanX + 4, lineY, contentFg, spanBg);
                        lineY += ch;
                    }
                }
            }
        }

        // Status bar
        int maxChars = formWidth / cw;
        string status;
        if (_searching)
        {
            status = $" Find: {_searchInput}\u2502  (Enter=Search  Esc=Cancel)";
        }
        else if (_editMode.IsActive())
        {
            status = _editMode.GetStatusText();
        }
        else
        {
            string cellRef = _grid.GetCellReference(selRow, selCol);
            string value = _grid.GetCellValue(selRow, selCol);
            string valueDisplay = string.IsNullOrEmpty(value) ? "" : $" = {value}";

            // F1 toggle: show resolved/expanded value
            string resolvedDisplay = "";
            if (_showResolved && !string.IsNullOrEmpty(value))
            {
                string resolved;
                string? inlineResult = _grid.ResolveInline(selRow, selCol);
                if (inlineResult != null && CellPrefix.IsCommand(inlineResult))
                {
                    resolved = _processManager.GetOutput(selRow, selCol) ?? "[running...]";
                }
                else if (inlineResult != null)
                {
                    resolved = CellPrefix.ExpandCellReferences(inlineResult, _grid);
                }
                else
                {
                    resolved = CellPrefix.ExpandCellReferences(value, _grid);
                }
                if (resolved != value)
                    resolvedDisplay = $" \u2192 {resolved}";
            }

            double? sum = _grid.GetColumnSum(selCol);
            string colName = GridManager.GetColumnName(selCol);
            string sumDisplay = sum.HasValue ? $"  \u03a3{colName} = {sum.Value}" : "";
            double? product = _grid.GetRowProduct(selRow);
            string productDisplay = product.HasValue ? $"  \u03a0{(selRow + 1)} = {product.Value}" : "";
            string searchDisplay = _searchTerm != null
                ? $"  \U0001f50d\"{_searchTerm}\" {(_searchMatches.Count > 0 ? $"{_searchMatchIndex + 1}/{_searchMatches.Count}" : "no matches")}"
                : "";
            string f1Label = _showResolved ? "F1: Raw" : "F1: Resolve";
            status = $" {cellRef}{valueDisplay}{resolvedDisplay}{sumDisplay}{productDisplay}{searchDisplay}  |  {f1Label}  F2: Edit  Ctrl+S: Save  Ctrl+Q: Quit";
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

    /// <summary>
    /// Strips common markdown syntax for cleaner display in span windows.
    /// </summary>
    private static string StripMarkdown(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new System.Text.StringBuilder(text.Length);
        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            // Strip heading markers: ### heading → heading
            if (line.StartsWith('#'))
            {
                int j = 0;
                while (j < line.Length && line[j] == '#') j++;
                line = line[j..].TrimStart();
            }
            // Strip horizontal rules
            if (line.Length >= 3 && line.All(c => c == '-' || c == '*' || c == '_' || c == ' '))
            {
                int dashes = line.Count(c => c == '-' || c == '*' || c == '_');
                if (dashes >= 3) { sb.AppendLine(""); continue; }
            }
            // Strip bold/italic markers
            line = line.Replace("***", "").Replace("**", "").Replace("__", "");
            // Strip inline code backticks
            line = line.Replace("`", "");
            // Strip bullet markers: - item or * item → item
            string trimmed = line.TrimStart();
            if (trimmed.Length > 1 && (trimmed[0] == '-' || trimmed[0] == '*') && trimmed[1] == ' ')
                line = "  " + trimmed[2..];
            sb.AppendLine(line);
        }
        return sb.ToString().TrimEnd();
    }
    
    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        if (_searching)
        {
            bool handled = true;
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    CommitSearch();
                    break;
                case Keys.Escape:
                    CancelSearch();
                    break;
                case Keys.Back:
                    if (_searchInput.Length > 0)
                        _searchInput = _searchInput[..^1];
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

        if (_editMode.IsActive())
        {
            bool handled = _editMode.HandleKeyEventWindows(e);
            if (handled)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                Invalidate();
            }
            return;
        }

        // While search results active, only allow navigation/clear
        if (_searchTerm != null)
        {
            bool searchHandled = true;
            if (e.KeyCode == Keys.Enter && e.Shift && _searchMatches.Count > 0)
            {
                _searchMatchIndex = (_searchMatchIndex - 1 + _searchMatches.Count) % _searchMatches.Count;
                _grid.SelectCell(_searchMatches[_searchMatchIndex].row, _searchMatches[_searchMatchIndex].col);
            }
            else if (e.KeyCode == Keys.Enter && _searchMatches.Count > 0)
            {
                _searchMatchIndex = (_searchMatchIndex + 1) % _searchMatches.Count;
                _grid.SelectCell(_searchMatches[_searchMatchIndex].row, _searchMatches[_searchMatchIndex].col);
            }
            else if (e.KeyCode == Keys.Escape)
            {
                _searchTerm = null;
                _searchMatches.Clear();
                _searchMatchIndex = -1;
            }
            else if (e.Control && e.KeyCode == Keys.Q)
            {
                Close();
                return;
            }
            else
            {
                searchHandled = false;
            }
            if (searchHandled)
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
                    _grid.DeleteSelectedRow();
                    break;
                case Keys.C:
                    if (e.Shift)
                    {
                        // Ctrl+Shift+C: copy resolved/displayed content
                        var (curRow, curCol) = _grid.GetCurrentCell();
                        string raw = _grid.GetSelectedCellValue();
                        string display;
                        string? resolved = _grid.ResolveSelectedInline();
                        if (resolved != null && CellPrefix.IsCommand(resolved))
                        {
                            display = _processManager.GetOutput(curRow, curCol) ?? "[running...]";
                        }
                        else if (resolved != null)
                        {
                            display = CellPrefix.ExpandCellReferences(resolved, _grid);
                        }
                        else
                        {
                            // Normal cell — expand any {A1::C10} references
                            display = CellPrefix.ExpandCellReferences(raw, _grid);
                        }
                        _clipboard = display;
                        if (!string.IsNullOrEmpty(_clipboard))
                            Clipboard.SetText(_clipboard);
                        break;
                    }
                    // Ctrl+C: copy raw cell value
                    if (_selection.Count > 0)
                    {
                        var copyCells = new HashSet<(int row, int col)>(_selection) { _grid.GetCurrentCell() };
                        var sorted = copyCells.OrderBy(c => c.row).ThenBy(c => c.col).ToList();
                        _clipboard = string.Join(Environment.NewLine, sorted.Select(c => _grid.GetCellValue(c.row, c.col)));
                    }
                    else
                        _clipboard = _grid.GetSelectedCellValue();
                    if (!string.IsNullOrEmpty(_clipboard))
                        Clipboard.SetText(_clipboard);
                    break;
                case Keys.X:
                    if (_selection.Count > 0)
                    {
                        var cutCells = new HashSet<(int row, int col)>(_selection) { _grid.GetCurrentCell() };
                        var sorted = cutCells.OrderBy(c => c.row).ThenBy(c => c.col).ToList();
                        _clipboard = string.Join(Environment.NewLine, sorted.Select(c => _grid.GetCellValue(c.row, c.col)));
                        foreach (var (r, c) in sorted) _grid.SetCellValue(r, c, "");
                        _selection.Clear();
                    }
                    else
                    {
                        _clipboard = _grid.GetSelectedCellValue();
                        _grid.SetSelectedCellValue("");
                    }
                    if (!string.IsNullOrEmpty(_clipboard))
                        Clipboard.SetText(_clipboard);
                    break;
                case Keys.V:
                    string paste = Clipboard.ContainsText() ? Clipboard.GetText() : _clipboard;
                    string[] lines = paste.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                    {
                        var (curRow, curCol) = _grid.GetCurrentCell();
                        for (int i = 0; i < lines.Length; i++)
                        {
                            int targetRow = curRow + i;
                            if (targetRow >= _grid.RowCount) break;
                            _grid.SetCellValue(targetRow, curCol, lines[i]);
                        }
                    }
                    break;
                case Keys.O: _grid.ShiftSelectedRowDown(); break;
                case Keys.P: _grid.ShiftSelectedRowUp(); break;
                case Keys.S: SaveFile(); break;
                case Keys.F: EnterSearchMode(); break;
                default: handled2 = false; break;
            }
        }
        else if (e.Shift)
        {
            // Shift+Arrow: extend multi-selection
            _selection.Add(_grid.GetCurrentCell());
            switch (e.KeyCode)
            {
                case Keys.Up: _grid.MoveUp(); break;
                case Keys.Down: _grid.MoveDown(); break;
                case Keys.Left: _grid.MoveLeft(); break;
                case Keys.Right: _grid.MoveRight(); break;
                default: handled2 = false; break;
            }
            if (handled2) _selection.Add(_grid.GetCurrentCell());
        }
        else
        {
            switch (e.KeyCode)
            {
                case Keys.Up:
                    _grid.MoveUp();
                    _selection.Clear();
                    break;
                case Keys.Down:
                    _grid.MoveDown();
                    _selection.Clear();
                    break;
                case Keys.Left:
                    _grid.MoveLeft();
                    _selection.Clear();
                    break;
                case Keys.Right:
                    _grid.MoveRight();
                    _selection.Clear();
                    break;
                case Keys.Enter:
                    {
                        var (curRow, curCol) = _grid.GetCurrentCell();
                        _extensionManager.ReactivateCell(curRow, curCol);
                        // If current cell is an inline ref to a command, re-run it
                        if (TryRerunInlineCommand(curRow, curCol))
                            break;
                        OpenAllSelected();
                    }
                    break;
                case Keys.F1:
                    _showResolved = !_showResolved;
                    break;
                case Keys.F2:
                    _editMode.Enter();
                    break;
                case Keys.F3:
                    RebuildGrid();
                    break;
                case Keys.F4:
                    if (_columnWidth > MinColumnWidth)
                    {
                        _columnWidth -= ColumnWidthStep;
                        RebuildGrid();
                    }
                    break;
                case Keys.F5:
                    if (_columnWidth < MaxColumnWidth)
                    {
                        _columnWidth += ColumnWidthStep;
                        RebuildGrid();
                    }
                    break;
                case Keys.Back:
                    {
                        var val = _grid.GetSelectedCellValue();
                        if (val.Length > 0)
                            _grid.SetSelectedCellValue(val[..^1]);
                    }
                    break;
                case Keys.Delete:
                    if (_selection.Count > 0)
                    {
                        foreach (var (r, c) in _selection) _grid.SetCellValue(r, c, "");
                        _grid.SetSelectedCellValue("");
                        _selection.Clear();
                    }
                    else
                    {
                        _grid.SetSelectedCellValue("");
                    }
                    break;
                case Keys.Tab:
                    {
                        var (curRow, curCol) = _grid.GetCurrentCell();
                        if (curCol < _grid.ColumnCount - 1) _grid.SelectCell(curRow, curCol + 1);
                        else if (curRow < _grid.RowCount - 1) _grid.SelectCell(curRow + 1, 0);
                    }
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

    private void EnterSearchMode()
    {
        _searching = true;
        _searchInput = "";
    }

    private void CommitSearch()
    {
        _searching = false;
        if (string.IsNullOrEmpty(_searchInput))
        {
            _searchTerm = null;
            _searchMatches.Clear();
            _searchMatchIndex = -1;
            return;
        }

        _searchTerm = _searchInput;
        _searchMatches.Clear();
        for (int r = 0; r < _grid.RowCount; r++)
            for (int c = 0; c < _grid.ColumnCount; c++)
                if (_grid.GetCellValue(r, c).Contains(_searchInput, StringComparison.OrdinalIgnoreCase))
                    _searchMatches.Add((r, c));

        if (_searchMatches.Count > 0)
        {
            _searchMatchIndex = 0;
            _grid.SelectCell(_searchMatches[0].row, _searchMatches[0].col);
        }
        else
        {
            _searchMatchIndex = -1;
        }
    }

    private void CancelSearch()
    {
        _searching = false;
        _searchInput = "";
    }

    private void OnFormKeyPress(object? sender, KeyPressEventArgs e)
    {
        if (e.KeyChar >= 32 && e.KeyChar <= 126)
        {
            if (_searching)
            {
                _searchInput += e.KeyChar;
            }
            else if (_editMode.IsActive())
            {
                _editMode.HandleCharInput(e.KeyChar);
            }
            else
            {
                _grid.AppendToSelectedCell(e.KeyChar);
            }
            e.Handled = true;
            Invalidate();
        }
    }

    private (int row, int col)? HitTestCell(int mouseX, int mouseY)
    {
        int headerRows = 2;
        int row = (mouseY / _charHeight) - headerRows;
        if (row < 0 || row >= _grid.RowCount) return null;

        int[] colWidths = GetColumnWidths();
        int x = RowHeaderWidth * _charWidth;
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            int colPx = colWidths[c] * _charWidth;
            if (mouseX >= x && mouseX < x + colPx) return (row, c);
            x += colPx;
        }
        return null;
    }

    private void SelectRectangle(int anchorRow, int anchorCol, int endRow, int endCol)
    {
        _selection.Clear();
        int r1 = Math.Min(anchorRow, endRow);
        int r2 = Math.Max(anchorRow, endRow);
        int c1 = Math.Min(anchorCol, endCol);
        int c2 = Math.Max(anchorCol, endCol);
        for (int r = r1; r <= r2; r++)
            for (int c = c1; c <= c2; c++)
                _selection.Add((r, c));
    }

    private void OnFormMouseDown(object? sender, MouseEventArgs e)
    {
        var cell = HitTestCell(e.X, e.Y);
        if (cell is null) return;
        var (row, col) = cell.Value;

        if (ModifierKeys.HasFlag(Keys.Control))
        {
            if (!_selection.Remove((row, col)))
                _selection.Add((row, col));
        }
        else
        {
            _selection.Clear();
            _dragging = true;
            _dragAnchorRow = row;
            _dragAnchorCol = col;
        }

        _grid.SelectCell(row, col);
        Invalidate();
    }

    private void OnFormMouseMove(object? sender, MouseEventArgs e)
    {
        if (!_dragging) return;
        var cell = HitTestCell(e.X, e.Y);
        if (cell is null) return;
        var (row, col) = cell.Value;

        _grid.SelectCell(row, col);
        SelectRectangle(_dragAnchorRow, _dragAnchorCol, row, col);
        Invalidate();
    }

    private void OnFormMouseUp(object? sender, MouseEventArgs e)
    {
        _dragging = false;
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

    public override void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        base.OnFormClosing(sender, e);
        _processManager.StopAll();
        _grid.SaveToCsv(AutoSavePath);
        _trayIcon.Visible = false;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _loopManager.Dispose();
            _inlineRefreshTimer?.Dispose();
            _autoSaveTimer?.Dispose();
            _csvReloadTimer?.Dispose();
            _processManager.Dispose();
            _extensionManager.Dispose();
            _monoFont.Dispose();
            _trayIcon.Dispose();
        }
        base.Dispose(disposing);
    }
}
