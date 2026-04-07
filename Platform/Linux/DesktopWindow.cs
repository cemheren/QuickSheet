using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using static ExcelConsole.Platform.Linux.X11Methods;

namespace ExcelConsole.Platform.Linux;

/// <summary>
/// X11 window that renders the spreadsheet grid as a desktop background.
/// Uses _NET_WM_WINDOW_TYPE_DESKTOP to sit below all other windows.
/// Rendering via Xlib drawing primitives + Xft for anti-aliased text.
/// </summary>
internal class DesktopWindow : IDisposable
{
    private readonly GridManager _grid;
    private int _selectedRow;
    private int _selectedCol;
    private readonly HashSet<(int row, int col)> _selection = new();
    private string _clipboard = "";
    private string? _loadedFile;

    private IntPtr _display;
    private IntPtr _window;
    private IntPtr _gc;
    private IntPtr _xftDraw;
    private IntPtr _xftFont;
    private IntPtr _visual;
    private IntPtr _colormap;
    private int _screen;

    private int _charWidth;
    private int _charHeight;
    private int _fontAscent;
    private int _screenWidth;
    private int _screenHeight;

    private readonly int _colWidth;
    private const int RowHeaderWidth = 4;
    private const string FontName = "monospace:size=14";

    private bool _running;
    private bool _isNativeX11;
    private System.Threading.Timer? _autoSaveTimer;
    private readonly object _saveLock = new();

    // Double-click tracking
    private ulong _lastClickTime;
    private int _lastClickRow = -1;
    private int _lastClickCol = -1;
    private const ulong DoubleClickMs = 400;

    // Transparency
    private bool _hasArgbVisual;
    private int _bgAlpha = 204; // ~80% opacity (0-255)

    // X11 atoms
    private IntPtr _atomWmDeleteWindow;
    private IntPtr _atomClipboard;
    private IntPtr _atomTargets;
    private IntPtr _atomUtf8String;
    private IntPtr _atomString;

    private static readonly string StateDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelConsole");
    private static readonly string AutoSavePath = Path.Combine(StateDir, "autosave.csv");

    public DesktopWindow(IntPtr display, string? csvPath)
    {
        _display = display;
        _screen = XDefaultScreen(display);
        _visual = XDefaultVisual(display, _screen);
        _colormap = XDefaultColormap(display, _screen);
        _screenWidth = XDisplayWidth(display, _screen);
        _screenHeight = XDisplayHeight(display, _screen);

        string sessionType = Environment.GetEnvironmentVariable("XDG_SESSION_TYPE") ?? "";
        _isNativeX11 = sessionType.Equals("x11", StringComparison.OrdinalIgnoreCase)
                     || string.IsNullOrEmpty(sessionType);

        // Open font and measure character size
        _xftFont = XftFontOpenName(display, _screen, FontName);
        if (_xftFont == IntPtr.Zero)
        {
            _xftFont = XftFontOpenName(display, _screen, "monospace:size=12");
        }
        MeasureFont();

        int availableWidth = _screenWidth / _charWidth;
        int availableHeight = _screenHeight / _charHeight - 3;
        _grid = new GridManager(availableWidth, availableHeight);

        int usableChars = availableWidth - RowHeaderWidth;
        _colWidth = _grid.ColumnCount > 0 ? usableChars / _grid.ColumnCount : 20;

        // Create the window (try 32-bit ARGB visual for transparency)
        IntPtr root = XDefaultRootWindow(display);
        if (XMatchVisualInfo(display, _screen, 32, TrueColor, out var vinfo) != 0)
        {
            _visual = vinfo.visual;
            _colormap = XCreateColormap(display, root, _visual, AllocNone);
            _hasArgbVisual = true;

            var attrs = new XSetWindowAttributes
            {
                background_pixel = 0, // fully transparent black
                border_pixel = 0,
                colormap = _colormap
            };
            _window = XCreateWindow(display, root,
                0, 0, (uint)_screenWidth, (uint)_screenHeight, 0,
                32, InputOutput, _visual,
                CWBackPixel | CWBorderPixel | CWColormap, ref attrs);
        }
        else
        {
            _hasArgbVisual = false;
            _window = XCreateSimpleWindow(display, root,
                0, 0, (uint)_screenWidth, (uint)_screenHeight,
                0, 0, 0);
        }

        XStoreName(display, _window, "QuickSheet");

        // Create graphics context
        _gc = XCreateGC(display, _window, 0, IntPtr.Zero);

        // Subscribe to events
        XSelectInput(display, _window,
            KeyPressMask | ButtonPressMask | ExposureMask |
            StructureNotifyMask | FocusChangeMask);

        // Intern atoms
        _atomWmDeleteWindow = XInternAtom(display, "WM_DELETE_WINDOW", false);
        _atomClipboard = XInternAtom(display, "CLIPBOARD", false);
        _atomTargets = XInternAtom(display, "TARGETS", false);
        _atomUtf8String = XInternAtom(display, "UTF8_STRING", false);
        _atomString = XInternAtom(display, "STRING", false);

        // Handle WM close
        IntPtr[] protocols = [_atomWmDeleteWindow];
        SetWmProtocols(display, _window, protocols);

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

        PopulateDesktopFiles();

        // Autosave every 5 seconds
        Directory.CreateDirectory(StateDir);
        _autoSaveTimer = new System.Threading.Timer(_ =>
        {
            lock (_saveLock)
            {
                try
                {
                    string tmp = AutoSavePath + ".tmp";
                    _grid.SaveToCsv(tmp);
                    File.Move(tmp, AutoSavePath, overwrite: true);
                }
                catch { }
            }
        }, null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(5));
    }

    private void SetWmProtocols(IntPtr display, IntPtr window, IntPtr[] protocols)
    {
        IntPtr wmProtocols = XInternAtom(display, "WM_PROTOCOLS", false);
        IntPtr atomAtom = XInternAtom(display, "ATOM", false);
        XChangeProperty(display, window, wmProtocols, atomAtom,
            32, PropModeReplace, protocols, protocols.Length);
    }

    private void MeasureFont()
    {
        // Read font metrics directly from the XftFont struct
        // Layout: int ascent, int descent, int height, int max_advance_width
        int ascent = Marshal.ReadInt32(_xftFont, 0);
        int descent = Marshal.ReadInt32(_xftFont, 4);
        int height = Marshal.ReadInt32(_xftFont, 8);
        int maxAdvance = Marshal.ReadInt32(_xftFont, 12);

        _fontAscent = ascent;
        _charHeight = height;
        if (_charHeight < 1) _charHeight = ascent + descent;
        if (_charHeight < 1) _charHeight = 18;

        // Measure character width with a sample string
        byte[] sample = "MMMMMMMMMM"u8.ToArray();
        XftTextExtentsUtf8(_display, _xftFont, sample, sample.Length, out XGlyphInfo extents);
        _charWidth = (int)Math.Ceiling(extents.xOff / 10.0);
        if (_charWidth < 1) _charWidth = maxAdvance > 0 ? maxAdvance : 8;
    }

    // ── Desktop mode ─────────────────────────────────────────────────

    public void EnterDesktopMode()
    {
        IntPtr atomAtom = XInternAtom(_display, "ATOM", false);

        // Remove window decorations via Motif hints
        IntPtr motifHints = XInternAtom(_display, "_MOTIF_WM_HINTS", false);
        long[] mwmHints = [2, 0, 0, 0, 0]; // flags=MWM_HINTS_DECORATIONS, decorations=0
        IntPtr mwmPtr = Marshal.AllocHGlobal(mwmHints.Length * 8);
        Marshal.Copy(mwmHints, 0, mwmPtr, mwmHints.Length);
        XChangeProperty(_display, _window, motifHints, motifHints,
            32, PropModeReplace, mwmPtr, mwmHints.Length);
        Marshal.FreeHGlobal(mwmPtr);

        if (_isNativeX11)
        {
            // On native X11, set window type to desktop — WM treats it as the desktop
            IntPtr wmWindowType = XInternAtom(_display, "_NET_WM_WINDOW_TYPE", false);
            IntPtr wmWindowTypeDesktop = XInternAtom(_display, "_NET_WM_WINDOW_TYPE_DESKTOP", false);
            IntPtr[] typeValue = [wmWindowTypeDesktop];
            XChangeProperty(_display, _window, wmWindowType, atomAtom,
                32, PropModeReplace, typeValue, 1);
        }

        // State: below other windows, sticky across workspaces, skip taskbar/pager
        IntPtr wmState = XInternAtom(_display, "_NET_WM_STATE", false);
        IntPtr stateBelow = XInternAtom(_display, "_NET_WM_STATE_BELOW", false);
        IntPtr stateSticky = XInternAtom(_display, "_NET_WM_STATE_STICKY", false);
        IntPtr stateSkipTaskbar = XInternAtom(_display, "_NET_WM_STATE_SKIP_TASKBAR", false);
        IntPtr stateSkipPager = XInternAtom(_display, "_NET_WM_STATE_SKIP_PAGER", false);
        IntPtr stateMaxH = XInternAtom(_display, "_NET_WM_STATE_MAXIMIZED_HORZ", false);
        IntPtr stateMaxV = XInternAtom(_display, "_NET_WM_STATE_MAXIMIZED_VERT", false);

        IntPtr[] states = [stateBelow, stateSticky, stateSkipTaskbar, stateSkipPager, stateMaxH, stateMaxV];
        XChangeProperty(_display, _window, wmState, atomAtom,
            32, PropModeReplace, states, states.Length);
    }

    // ── Desktop files ────────────────────────────────────────────────

    private void PopulateDesktopFiles()
    {
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (string.IsNullOrEmpty(desktopPath) || !Directory.Exists(desktopPath))
        {
            // Fallback to XDG desktop dir
            string? xdgDesktop = Environment.GetEnvironmentVariable("XDG_DESKTOP_DIR");
            if (xdgDesktop is not null && Directory.Exists(xdgDesktop))
                desktopPath = xdgDesktop;
            else
            {
                string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                desktopPath = Path.Combine(home, "Desktop");
                if (!Directory.Exists(desktopPath)) return;
            }
        }

        var entries = Directory.GetFileSystemEntries(desktopPath)
            .OrderBy(e => Directory.Exists(e) ? 0 : 1)
            .ThenBy(Path.GetFileName)
            .ToArray();

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
            Process.Start(new ProcessStartInfo("xdg-open", path) { UseShellExecute = false });
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
        string cmd = value[3..].Trim();
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
            string fullCmd = string.IsNullOrEmpty(args) ? exe : $"{exe} {args}";
            var (terminal, termArgs) = FindTerminal();
            if (!string.IsNullOrEmpty(terminal))
            {
                Process.Start(new ProcessStartInfo(terminal, $"{termArgs} {fullCmd}")
                {
                    UseShellExecute = false,
                    CreateNoWindow = false
                });
            }
            else
            {
                // Fallback: run detached with nohup so it doesn't share our terminal
                Process.Start(new ProcessStartInfo("/bin/sh", $"-c \"nohup {fullCmd} &\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = false
                });
            }
        }
        catch { }
    }

    private static (string terminal, string launchArg) FindTerminal()
    {
        // (terminal, argument to launch a command)
        (string, string)[] terminals =
        [
            ("ptyxis", "--"),
            ("gnome-terminal", "--"),
            ("kgx", "-e"),
            ("konsole", "-e"),
            ("xfce4-terminal", "-e"),
            ("mate-terminal", "-e"),
            ("xterm", "-e")
        ];
        foreach (var (t, arg) in terminals)
        {
            try
            {
                var result = Process.Start(new ProcessStartInfo("which", t)
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false
                });
                result?.WaitForExit(500);
                if (result?.ExitCode == 0) return (t, arg);
            }
            catch { }
        }
        return ("", "");
    }

    private void OpenHyperlink()
    {
        string url = _grid.GetCellValue(_selectedRow, _selectedCol);
        if (!IsHyperlink(url)) return;
        try
        {
            Process.Start(new ProcessStartInfo("xdg-open", url) { UseShellExecute = false });
        }
        catch { }
    }

    private void OpenAllSelected()
    {
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
                    try { Process.Start(new ProcessStartInfo("xdg-open", path) { UseShellExecute = false }); opened = true; } catch { }
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
                    try { Process.Start(new ProcessStartInfo("xdg-open", val) { UseShellExecute = false }); opened = true; } catch { }
            }
        }

        if (!opened && _selectedRow < _grid.RowCount - 1)
            _selectedRow++;
    }

    // ── Event Loop ───────────────────────────────────────────────────

    public void Run()
    {
        EnterDesktopMode();
        XMapWindow(_display, _window);
        XSync(_display, false);

        // Create XftDraw now that window is mapped
        _xftDraw = XftDrawCreate(_display, _window, _visual, _colormap);

        _running = true;

        IntPtr eventPtr = Marshal.AllocHGlobal(XEventSize);
        try
        {
            while (_running)
            {
                XNextEvent(_display, eventPtr);
                int eventType = Marshal.ReadInt32(eventPtr);

                switch (eventType)
                {
                    case Expose:
                        Render();
                        break;

                    case KeyPress:
                        var keyEvent = Marshal.PtrToStructure<XKeyEvent>(eventPtr);
                        HandleKeyPress(ref keyEvent);
                        Render();
                        break;

                    case ButtonPress:
                        var buttonEvent = Marshal.PtrToStructure<XButtonEvent>(eventPtr);
                        HandleButtonPress(ref buttonEvent);
                        Render();
                        break;

                    case ConfigureNotify:
                        var configEvent = Marshal.PtrToStructure<XConfigureEvent>(eventPtr);
                        if (configEvent.width != _screenWidth || configEvent.height != _screenHeight)
                        {
                            _screenWidth = configEvent.width;
                            _screenHeight = configEvent.height;
                            if (_xftDraw != IntPtr.Zero)
                                XftDrawDestroy(_xftDraw);
                            _xftDraw = XftDrawCreate(_display, _window, _visual, _colormap);
                            Render();
                        }
                        break;

                    case ClientMessage:
                        var clientMsg = Marshal.PtrToStructure<XClientMessageEvent>(eventPtr);
                        if (clientMsg.data0 == (long)_atomWmDeleteWindow)
                            _running = false;
                        break;

                    case SelectionRequest:
                        var selReq = Marshal.PtrToStructure<XSelectionRequestEvent>(eventPtr);
                        HandleSelectionRequest(ref selReq);
                        break;

                    case SelectionNotify:
                        var selNotify = Marshal.PtrToStructure<XSelectionEvent>(eventPtr);
                        HandleSelectionNotify(ref selNotify);
                        break;
                }
            }
        }
        finally
        {
            Marshal.FreeHGlobal(eventPtr);
        }
    }

    public void Stop() => _running = false;

    // ── Rendering ────────────────────────────────────────────────────

    private void Render()
    {
        int cw = _charWidth;
        int ch = _charHeight;
        int[] colWidths = GetColumnWidths();

        // Clear background (semi-transparent if ARGB visual available)
        SetGCColor(0, 0, 0, _hasArgbVisual ? _bgAlpha : 255);
        XFillRectangle(_display, _window, _gc, 0, 0, (uint)_screenWidth, (uint)_screenHeight);

        int y = 0;

        // Column headers
        int x = RowHeaderWidth * cw;
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            int w = colWidths[c];
            string header = GridManager.GetColumnName(c).PadRight(w);
            if (c == _selectedCol)
                DrawTextWithBg(header, x, y, 255, 255, 255, 64, 64, 64);
            else
                DrawTextWithBg(header, x, y, 255, 255, 255, 0, 0, 0);
            x += w * cw;
        }
        y += ch;

        // Underline
        string underline = new string('-', RowHeaderWidth);
        for (int c = 0; c < _grid.ColumnCount; c++)
            underline += new string('-', colWidths[c]);
        DrawTextWithBg(underline, 0, y, 255, 255, 255, 0, 0, 0);
        y += ch;

        // Data rows
        int statusY = _screenHeight - ch;
        for (int r = 0; r < _grid.RowCount && y + ch <= statusY; r++)
        {
            x = 0;
            string rowNum = (r + 1).ToString().PadLeft(RowHeaderWidth - 1) + " ";
            if (r == _selectedRow)
                DrawTextWithBg(rowNum, x, y, 255, 255, 255, 64, 64, 64);
            else
                DrawTextWithBg(rowNum, x, y, 200, 200, 200, 15, 15, 15);

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

                int bgR, bgG, bgB, fgR, fgG, fgB;

                if (isCursor) { bgR = 64; bgG = 64; bgB = 64; }
                else if (isMultiSel) { bgR = 50; bgG = 50; bgB = 80; }
                else if (isFile) { bgR = 0; bgG = 40; bgB = 60; }
                else if (isLink) { bgR = 40; bgG = 0; bgB = 60; }
                else if (isCmd) { bgR = 40; bgG = 40; bgB = 0; }
                else { bgR = 15; bgG = 15; bgB = 15; }

                if (isFile) { fgR = 100; fgG = 200; fgB = 255; }
                else if (isLink) { fgR = 180; fgG = 140; fgB = 255; }
                else if (isCmd) { fgR = 255; fgG = 220; fgB = 100; }
                else { fgR = 255; fgG = 255; fgB = 255; }

                DrawTextWithBg(display, x, y, fgR, fgG, fgB, bgR, bgG, bgB);
                x += w * cw;
            }
            // Subtle horizontal grid line at the bottom of each row
            SetGCColor(40, 40, 40, _hasArgbVisual ? _bgAlpha : 255);
            XDrawLine(_display, _window, _gc, 0, y + ch - 1, _screenWidth, y + ch - 1);
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
        int maxChars = _screenWidth / cw;
        status = status.PadRight(maxChars);

        // White background status bar (fully opaque)
        DrawTextWithBg(status, 0, statusY, 0, 0, 0, 255, 255, 255, 255);

        XFlush(_display);
    }

    private int[] GetColumnWidths()
    {
        var widths = new int[_grid.ColumnCount];
        for (int c = 0; c < _grid.ColumnCount; c++)
            widths[c] = _colWidth;
        return widths;
    }

    private void SetGCColor(int r, int g, int b, int a = 255)
    {
        ulong pixel = _hasArgbVisual
            ? (ulong)(a << 24 | r << 16 | g << 8 | b)
            : (ulong)(r << 16 | g << 8 | b);
        XSetForeground(_display, _gc, pixel);
    }

    private void DrawTextWithBg(string text, int x, int y, int fgR, int fgG, int fgB, int bgR, int bgG, int bgB, int bgA = -1)
    {
        int w = text.Length * _charWidth;

        // Draw background rectangle (use per-cell alpha, or default _bgAlpha for normal cells)
        int alpha = bgA >= 0 ? bgA : (_hasArgbVisual ? _bgAlpha : 255);
        SetGCColor(bgR, bgG, bgB, alpha);
        XFillRectangle(_display, _window, _gc, x, y, (uint)w, (uint)_charHeight);

        // Draw text with Xft
        var renderColor = new XRenderColor
        {
            red = (ushort)(fgR * 257),
            green = (ushort)(fgG * 257),
            blue = (ushort)(fgB * 257),
            alpha = 0xFFFF
        };
        XftColorAllocValue(_display, _visual, _colormap, ref renderColor, out XftColor xftColor);

        byte[] utf8 = Encoding.UTF8.GetBytes(text);
        XftDrawStringUtf8(_xftDraw, ref xftColor, _xftFont, x, y + _fontAscent, utf8, utf8.Length);

        XftColorFree(_display, _visual, _colormap, ref xftColor);
    }

    // ── Input Handling ───────────────────────────────────────────────

    private void HandleKeyPress(ref XKeyEvent keyEvent)
    {
        ulong keysym = XLookupKeysym(ref keyEvent, 0);
        bool ctrl = (keyEvent.state & ControlMask) != 0;
        bool shift = (keyEvent.state & ShiftMask) != 0;

        if (ctrl)
        {
            switch (keysym)
            {
                case XK_q:
                    _running = false;
                    return;
                case XK_d:
                    _grid.DeleteRow(_selectedRow);
                    if (_selectedRow >= _grid.RowCount) _selectedRow = _grid.RowCount - 1;
                    return;
                case XK_c:
                    if (_selection.Count > 0)
                    {
                        var copyCells = new List<(int row, int col)>(_selection) { (_selectedRow, _selectedCol) };
                        copyCells.Sort((a, b) => a.row != b.row ? a.row.CompareTo(b.row) : a.col.CompareTo(b.col));
                        _clipboard = string.Join(Environment.NewLine, copyCells.Select(c => _grid.GetCellValue(c.row, c.col)));
                    }
                    else
                        _clipboard = _grid.GetCellValue(_selectedRow, _selectedCol);
                    SetX11Clipboard(_clipboard);
                    return;
                case XK_x:
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
                    SetX11Clipboard(_clipboard);
                    return;
                case XK_v:
                    string paste = GetX11Clipboard() ?? _clipboard;
                    string[] lines = paste.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        int targetRow = _selectedRow + i;
                        if (targetRow >= _grid.RowCount) break;
                        _grid.SetCellValue(targetRow, _selectedCol, lines[i]);
                    }
                    return;
                case XK_o:
                    _grid.ShiftRowsDown(_selectedRow);
                    return;
                case XK_p:
                    _grid.ShiftRowsUp(_selectedRow);
                    return;
                case XK_s:
                    SaveFile();
                    return;
            }
        }

        if (shift)
        {
            _selection.Add((_selectedRow, _selectedCol));
            bool handled = true;
            switch (keysym)
            {
                case XK_Up: if (_selectedRow > 0) _selectedRow--; break;
                case XK_Down: if (_selectedRow < _grid.RowCount - 1) _selectedRow++; break;
                case XK_Left: if (_selectedCol > 0) _selectedCol--; break;
                case XK_Right: if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++; break;
                default: handled = false; break;
            }
            if (handled)
            {
                _selection.Add((_selectedRow, _selectedCol));
                return;
            }
        }

        switch (keysym)
        {
            case XK_Up:
                if (_selectedRow > 0) _selectedRow--;
                _selection.Clear();
                break;
            case XK_Down:
                if (_selectedRow < _grid.RowCount - 1) _selectedRow++;
                _selection.Clear();
                break;
            case XK_Left:
                if (_selectedCol > 0) _selectedCol--;
                _selection.Clear();
                break;
            case XK_Right:
                if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                _selection.Clear();
                break;
            case XK_Return:
                OpenAllSelected();
                break;
            case XK_BackSpace:
                var val = _grid.GetCellValue(_selectedRow, _selectedCol);
                if (val.Length > 0) _grid.SetCellValue(_selectedRow, _selectedCol, val[..^1]);
                break;
            case XK_Delete:
                _grid.SetCellValue(_selectedRow, _selectedCol, "");
                break;
            case XK_Tab:
                if (_selectedCol < _grid.ColumnCount - 1) _selectedCol++;
                else if (_selectedRow < _grid.RowCount - 1) { _selectedCol = 0; _selectedRow++; }
                _selection.Clear();
                break;
            case XK_Escape:
                _selection.Clear();
                break;
            default:
                // Printable character input
                if (!ctrl)
                {
                    IntPtr buf = Marshal.AllocHGlobal(32);
                    try
                    {
                        int len = XLookupString(ref keyEvent, buf, 32, out _, IntPtr.Zero);
                        if (len > 0)
                        {
                            byte[] bytes = new byte[len];
                            Marshal.Copy(buf, bytes, 0, len);
                            string ch = Encoding.UTF8.GetString(bytes);
                            if (ch.Length > 0 && ch[0] >= 32 && ch[0] <= 126)
                            {
                                var cur = _grid.GetCellValue(_selectedRow, _selectedCol);
                                _grid.SetCellValue(_selectedRow, _selectedCol, cur + ch);
                            }
                        }
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(buf);
                    }
                }
                break;
        }
    }

    private void HandleButtonPress(ref XButtonEvent buttonEvent)
    {
        if (buttonEvent.button == 1) // Left click
        {
            int headerRows = 2;
            int row = (buttonEvent.y / _charHeight) - headerRows;
            if (row < 0 || row >= _grid.RowCount) return;

            int[] colWidths = GetColumnWidths();
            int x = RowHeaderWidth * _charWidth;
            int col = -1;
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                int colPx = colWidths[c] * _charWidth;
                if (buttonEvent.x >= x && buttonEvent.x < x + colPx) { col = c; break; }
                x += colPx;
            }
            if (col < 0) return;

            // Double-click detection
            bool isDoubleClick = row == _lastClickRow && col == _lastClickCol
                && buttonEvent.time - _lastClickTime < DoubleClickMs;
            _lastClickTime = buttonEvent.time;
            _lastClickRow = row;
            _lastClickCol = col;

            if (isDoubleClick)
            {
                _selectedRow = row;
                _selectedCol = col;
                OpenAllSelected();
                return;
            }

            if ((buttonEvent.state & ControlMask) != 0)
            {
                if (!_selection.Remove((row, col)))
                    _selection.Add((row, col));
            }
            else
            {
                _selection.Clear();
            }

            _selectedRow = row;
            _selectedCol = col;
        }
        else if (buttonEvent.button == 3) // Right click
        {
            OpenAllSelected();
        }
    }

    // ── Clipboard ────────────────────────────────────────────────────

    private void SetX11Clipboard(string text)
    {
        _clipboard = text;
        XSetSelectionOwner(_display, _atomClipboard, _window, 0);
    }

    private string? GetX11Clipboard()
    {
        IntPtr owner = XGetSelectionOwner(_display, _atomClipboard);
        if (owner == _window)
            return _clipboard;

        // Request clipboard from external owner
        IntPtr clipProp = XInternAtom(_display, "QUICKSHEET_CLIP", false);
        XConvertSelection(_display, _atomClipboard, _atomUtf8String, clipProp, _window, 0);
        XFlush(_display);

        // Wait for SelectionNotify (with timeout)
        IntPtr eventPtr = Marshal.AllocHGlobal(XEventSize);
        try
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(500);
            while (DateTime.UtcNow < deadline)
            {
                if (XPending(_display) > 0)
                {
                    XNextEvent(_display, eventPtr);
                    int eventType = Marshal.ReadInt32(eventPtr);
                    if (eventType == SelectionNotify)
                    {
                        var selEvent = Marshal.PtrToStructure<XSelectionEvent>(eventPtr);
                        if (selEvent.property != IntPtr.Zero)
                        {
                            string? result = ReadPropertyString(selEvent.property);
                            return result;
                        }
                        return null;
                    }
                }
                else
                {
                    Thread.Sleep(10);
                }
            }
        }
        finally
        {
            Marshal.FreeHGlobal(eventPtr);
        }
        return null;
    }

    private string? ReadPropertyString(IntPtr property)
    {
        int result = XGetWindowProperty(_display, _window, property,
            0, 1024, true, _atomUtf8String,
            out _, out _, out ulong nitems, out _, out IntPtr data);

        if (result != 0 || data == IntPtr.Zero || nitems == 0)
        {
            if (data != IntPtr.Zero) XFree(data);
            return null;
        }

        byte[] buffer = new byte[nitems];
        Marshal.Copy(data, buffer, 0, (int)nitems);
        XFree(data);
        return Encoding.UTF8.GetString(buffer);
    }

    private void HandleSelectionRequest(ref XSelectionRequestEvent req)
    {
        if (req.target == _atomTargets)
        {
            IntPtr[] targets = [_atomUtf8String, _atomString];
            IntPtr atomAtom = XInternAtom(_display, "ATOM", false);
            XChangeProperty(_display, req.requestor, req.property, atomAtom,
                32, PropModeReplace, targets, targets.Length);
        }
        else if (req.target == _atomUtf8String || req.target == _atomString)
        {
            byte[] data = Encoding.UTF8.GetBytes(_clipboard);
            IntPtr dataPtr = Marshal.AllocHGlobal(data.Length);
            Marshal.Copy(data, 0, dataPtr, data.Length);
            XChangeProperty(_display, req.requestor, req.property, _atomUtf8String,
                8, PropModeReplace, dataPtr, data.Length);
            Marshal.FreeHGlobal(dataPtr);
        }

        // Send SelectionNotify response
        var response = new XSelectionEvent
        {
            type = SelectionNotify,
            requestor = req.requestor,
            selection = req.selection,
            target = req.target,
            property = req.property,
            time = req.time
        };
        IntPtr responsePtr = Marshal.AllocHGlobal(XEventSize);
        Marshal.StructureToPtr(response, responsePtr, false);
        XSendEvent(_display, req.requestor, false, 0, responsePtr);
        Marshal.FreeHGlobal(responsePtr);
        XFlush(_display);
    }

    private void HandleSelectionNotify(ref XSelectionEvent selEvent)
    {
        // Handled inline in GetX11Clipboard
    }

    // ── File operations ──────────────────────────────────────────────

    private void SaveFile()
    {
        string filename = _loadedFile ?? "spreadsheet.csv";
        _grid.SaveToCsv(filename);
        _loadedFile = filename;
    }

    // ── Dispose ──────────────────────────────────────────────────────

    public void Dispose()
    {
        _autoSaveTimer?.Dispose();

        // Autosave on exit
        try
        {
            Directory.CreateDirectory(StateDir);
            _grid.SaveToCsv(AutoSavePath);
        }
        catch { }

        if (_xftDraw != IntPtr.Zero)
            XftDrawDestroy(_xftDraw);
        if (_xftFont != IntPtr.Zero)
            XftFontClose(_display, _xftFont);
        if (_gc != IntPtr.Zero)
            XFreeGC(_display, _gc);
        if (_window != IntPtr.Zero)
            XDestroyWindow(_display, _window);
    }
}
