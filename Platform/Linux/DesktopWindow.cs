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
    private GridManager _grid;
    private int _selectedRow;
    private int _selectedCol;
    private readonly HashSet<(int row, int col)> _selection = new();
    private string _clipboard = "";
    private string? _loadedFile;

    private bool _searching;
    private string _searchInput = "";
    private string? _searchTerm;
    private List<(int row, int col)> _searchMatches = new();
    private int _searchMatchIndex = -1;

    private readonly LinuxEditingMode _editMode;
    private bool _showResolved;
    private readonly InlineProcessManager _inlineProcesses = new();
    private readonly Extensions.ExtensionManager _extensionManager;
    private readonly LoopManager _loopManager;

    private IntPtr _display;
    private IntPtr _window;
    private IntPtr _gc;
    private IntPtr _xftDraw;
    private IntPtr _xftFont;
    private IntPtr _visual;
    private IntPtr _colormap;
    private IntPtr _backBuffer;
    private IntPtr _backBufferXftDraw;
    private IntPtr _renderDrawable;  // current draw target (backbuffer or window)
    private IntPtr _renderXftDraw;   // current xft draw target
    private int _screen;

    private int _charWidth;
    private int _charHeight;
    private int _fontAscent;
    private int _screenWidth;
    private int _screenHeight;

    private int _columnWidth = 20;
    private const int MinColumnWidth = 8;
    private const int MaxColumnWidth = 40;
    private const int ColumnWidthStep = 4;
    private const int RowHeaderWidth = 4;
    private const string FontName = "monospace:size=14";

    private bool _running;
    private bool _isNativeX11;
    private System.Threading.Timer? _autoSaveTimer;
    private readonly object _saveLock = new();
    private volatile bool _externalChangePending;

    // Double-click tracking
    private ulong _lastClickTime;
    private int _lastClickRow = -1;
    private int _lastClickCol = -1;
    private const ulong DoubleClickMs = 400;

    // Drag selection
    private bool _dragging;
    private int _dragAnchorRow;
    private int _dragAnchorCol;

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
        _grid = new GridManager(availableWidth, availableHeight, _columnWidth);
        _editMode = new LinuxEditingMode(_grid);
        _extensionManager = new Extensions.ExtensionManager(new LinuxExtensionEnvironment(), _grid);
        _loopManager = new LoopManager(_grid, (r, c) => ActivateCellAt(r, c));

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
            KeyPressMask | ButtonPressMask | ButtonReleaseMask | Button1MotionMask |
            ExposureMask | StructureNotifyMask | FocusChangeMask);

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

        // Autosave every 5 seconds; every 12th tick (60s) also re-merge the source CSV
        // to pick up external edits (e.g. OneDrive/Syncthing) the same way Windows does.
        Directory.CreateDirectory(StateDir);
        int reloadTick = 0;
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

                if (++reloadTick >= 12)
                {
                    reloadTick = 0;
                    try
                    {
                        string? path = _loadedFile;
                        if (path is not null && _grid.MergeFromCsv(path))
                            _externalChangePending = true;
                    }
                    catch { }
                }
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
            var psi = new ProcessStartInfo("xdg-open") { UseShellExecute = false };
            psi.ArgumentList.Add(path);
            Process.Start(psi);
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

    private static readonly HashSet<string> ShellCommands = new(StringComparer.OrdinalIgnoreCase)
        { "bash", "sh", "zsh", "fish", "csh", "tcsh", "ksh", "dash" };

    private static void RunCommand(string cellValue)
    {
        if (!IsCommand(cellValue)) return;
        var (exe, args) = ParseCommand(cellValue);
        if (string.IsNullOrEmpty(exe)) return;
        try
        {
            bool needsTerminal = ShellCommands.Contains(Path.GetFileName(exe));
            if (needsTerminal)
            {
                var (terminal, termArgs) = FindTerminal();
                if (!string.IsNullOrEmpty(terminal))
                {
                    var psi = new ProcessStartInfo(terminal) { UseShellExecute = false, CreateNoWindow = false };
                    foreach (var part in termArgs.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                        psi.ArgumentList.Add(part);
                    psi.ArgumentList.Add(exe);
                    if (!string.IsNullOrEmpty(args))
                    {
                        psi.ArgumentList.Add("-c");
                        psi.ArgumentList.Add($"{args}; exec {exe}");
                    }
                    Process.Start(psi);
                    return;
                }
            }

            // Run directly (GUI apps, or fallback)
            var direct = new ProcessStartInfo(exe) { UseShellExecute = false, CreateNoWindow = false };
            if (!string.IsNullOrEmpty(args))
                direct.Arguments = args;
            Process.Start(direct);
        }
        catch { }
    }

    private static (string terminal, string launchArg) FindTerminal()
    {
        // (terminal, argument to launch a command)
        // ptyxis uses -s (standalone) + -- to avoid D-Bus activation issues
        (string, string)[] terminals =
        [
            ("ptyxis", "-s --"),
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

    // ── Event Loop ───────────────────────────────────────────────────

    public void Run()
    {
        EnterDesktopMode();
        XMapWindow(_display, _window);
        XSync(_display, false);

        // Create XftDraw now that window is mapped
        _xftDraw = XftDrawCreate(_display, _window, _visual, _colormap);

        // Create double buffer to eliminate flicker
        CreateBackBuffer();

        _running = true;

        IntPtr eventPtr = Marshal.AllocHGlobal(XEventSize);
        try
        {
            while (_running)
            {
                // If no events are pending, poll inline-process output and idle briefly.
                // This gives sub-second update cadence for `i:` cells without a full event-driven wakeup.
                if (XPending(_display) == 0)
                {
                    _extensionManager.ScanGrid();
                    if (_inlineProcesses.HasAnyNewOutput() || _externalChangePending || _extensionManager.ConsumeHasChanges())
                    {
                        _externalChangePending = false;
                        Render();
                    }
                    System.Threading.Thread.Sleep(100);
                    continue;
                }

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

                    case ButtonRelease:
                        HandleButtonRelease();
                        break;

                    case MotionNotify:
                        var motionEvent = Marshal.PtrToStructure<XMotionEvent>(eventPtr);
                        HandleMotionNotify(ref motionEvent);
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
                            CreateBackBuffer();
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
        _renderDrawable = _backBuffer != IntPtr.Zero ? _backBuffer : _window;
        _renderXftDraw = _backBufferXftDraw != IntPtr.Zero ? _backBufferXftDraw : _xftDraw;

        int cw = _charWidth;
        int ch = _charHeight;
        int[] colWidths = GetColumnWidths();

        // Clear background (semi-transparent if ARGB visual available)
        SetGCColor(0, 0, 0, _hasArgbVisual ? _bgAlpha : 255);
        XFillRectangle(_display, _renderDrawable, _gc, 0, 0, (uint)_screenWidth, (uint)_screenHeight);

        // Pre-scan for inline cells with multi-cell visual spans (e.g. {A1::C5}-style refs).
        // Each spanning inline anchor produces an overlay drawn after the row loop;
        // cells covered by the span (other than the anchor) are skipped during row rendering.
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
                int endColIdx = Math.Min(c + sc - 1, _grid.ColumnCount - 1);
                int endRowIdx = Math.Min(r + sr - 1, _grid.RowCount - 1);

                string? resolved = _grid.ResolveInline(r, c);
                bool isCmdSpan = false;
                string content;
                if (resolved != null && CellPrefix.IsCommand(resolved))
                {
                    isCmdSpan = true;
                    string expandedCmd = CellPrefix.ExpandCellReferences(resolved, _grid);
                    int ptyCols = -1; // padding budget
                    for (int tc = c; tc <= endColIdx; tc++) ptyCols += colWidths[tc];
                    if (ptyCols < 20) ptyCols = 20;
                    int ptyRows = Math.Max(1, (endRowIdx - r + 1) - 1);
                    _inlineProcesses.EnsureRunning(r, c, expandedCmd, ptyCols, ptyRows);
                    content = _inlineProcesses.GetOutput(r, c) ?? "[running...]";
                }
                else if (resolved != null)
                {
                    content = CellPrefix.ExpandCellReferences(resolved, _grid);
                }
                else
                {
                    content = val;
                }

                inlineSpans.Add((r, c, sc, sr, content, isCmdSpan));

                // Anchor stays renderable; only cover the rest of the span.
                for (int wr = r; wr <= endRowIdx; wr++)
                    for (int wc = c; wc <= endColIdx; wc++)
                        if (wr != r || wc != c)
                            occludedCells.Add((wr, wc));
            }
        }

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

                // Skip cells covered by an inline-span overlay (drawn later).
                if (occludedCells.Contains((r, c)))
                {
                    x += w * cw;
                    continue;
                }

                bool isEditingThisCell = _editMode.IsActive() && _editMode.EditRow == r && _editMode.EditCol == c;
                string cellVal = _grid.GetCellValue(r, c);

                // Resolve single-cell inline references for in-cell display.
                string displayVal = cellVal;
                if (!isEditingThisCell && CellPrefix.IsInline(cellVal))
                {
                    string? resolvedCell = _grid.ResolveInline(r, c);
                    if (resolvedCell != null)
                    {
                        if (CellPrefix.IsCommand(resolvedCell))
                        {
                            string expandedCmd = CellPrefix.ExpandCellReferences(resolvedCell, _grid);
                            _inlineProcesses.EnsureRunning(r, c, expandedCmd);
                            displayVal = _inlineProcesses.GetOutput(r, c) ?? "[running...]";
                            int nl = displayVal.IndexOf('\n');
                            if (nl >= 0) displayVal = displayVal[..nl];
                        }
                        else
                        {
                            displayVal = CellPrefix.ExpandCellReferences(resolvedCell, _grid);
                        }
                    }
                }

                string display = isEditingThisCell
                    ? _editMode.GetCellDisplay(w)
                    : (displayVal.Length >= w ? displayVal[..w] : displayVal.PadRight(w));
                bool isCursor = r == _selectedRow && c == _selectedCol;
                bool isMultiSel = _selection.Contains((r, c));
                bool isSearchMatch = _searchTerm != null && _searchMatches.Contains((r, c));
                bool isFile = _grid.IsFileEntry(r, c);
                bool isLink = IsHyperlink(cellVal);
                bool isCmd = IsCommand(cellVal);
                bool isLoop = CellPrefix.IsLoop(cellVal);
                var extStatus = _extensionManager.GetCellStatus(r, c);

                int bgR, bgG, bgB, fgR, fgG, fgB;

                if (isCursor && isSearchMatch) { bgR = 0; bgG = 180; bgB = 0; }
                else if (isCursor) { bgR = 64; bgG = 64; bgB = 64; }
                else if (isMultiSel) { bgR = 50; bgG = 50; bgB = 80; }
                else if (isSearchMatch) { bgR = 80; bgG = 80; bgB = 0; }
                else if (isFile) { bgR = 0; bgG = 40; bgB = 60; }
                else if (isLink) { bgR = 40; bgG = 0; bgB = 60; }
                else if (isCmd) { bgR = 40; bgG = 40; bgB = 0; }
                else if (isLoop) { bgR = 0; bgG = 40; bgB = 40; }
                else if (extStatus == Extensions.ExtensionCellStatus.Error) { bgR = 50; bgG = 10; bgB = 10; }
                else if (extStatus == Extensions.ExtensionCellStatus.Running) { bgR = 10; bgG = 40; bgB = 10; }
                else { bgR = 15; bgG = 15; bgB = 15; }

                if (isFile) { fgR = 100; fgG = 200; fgB = 255; }
                else if (isLink) { fgR = 180; fgG = 140; fgB = 255; }
                else if (isCmd) { fgR = 255; fgG = 220; fgB = 100; }
                else if (isLoop) { fgR = 100; fgG = 220; fgB = 200; }
                else if (extStatus == Extensions.ExtensionCellStatus.Error) { fgR = 255; fgG = 80; fgB = 80; }
                else if (extStatus == Extensions.ExtensionCellStatus.Running) { fgR = 80; fgG = 255; fgB = 80; }
                else { fgR = 255; fgG = 255; fgB = 255; }

                DrawTextWithBg(display, x, y, fgR, fgG, fgB, bgR, bgG, bgB);
                x += w * cw;
            }
            // Subtle horizontal grid line at the bottom of each row
            SetGCColor(40, 40, 40, _hasArgbVisual ? _bgAlpha : 255);
            XDrawLine(_display, _renderDrawable, _gc, 0, y + ch - 1, _screenWidth, y + ch - 1);
            y += ch;
        }

        // Inline span overlays (drawn over the cell layer for spanning i: cells)
        const int headerRows = 2; // column header + underline
        foreach (var span in inlineSpans)
        {
            int spanX = RowHeaderWidth * cw;
            for (int sc = 0; sc < span.anchorCol && sc < _grid.ColumnCount; sc++)
                spanX += colWidths[sc] * cw;

            int spanY = headerRows * ch + span.anchorRow * ch;
            int endColIdx = Math.Min(span.anchorCol + span.spanCols - 1, _grid.ColumnCount - 1);
            int endRowIdx = Math.Min(span.anchorRow + span.spanRows - 1, _grid.RowCount - 1);

            int spanWidth = 0;
            for (int sc = span.anchorCol; sc <= endColIdx; sc++)
                spanWidth += colWidths[sc] * cw;
            int spanHeight = (endRowIdx - span.anchorRow + 1) * ch;

            if (spanY + spanHeight > statusY) spanHeight = statusY - spanY;
            if (spanWidth <= 0 || spanHeight <= 0) continue;

            bool hasCursor = _selectedRow >= span.anchorRow && _selectedRow <= endRowIdx
                          && _selectedCol >= span.anchorCol && _selectedCol <= endColIdx;

            int sBgR, sBgG, sBgB;
            if (span.isCmd) { sBgR = hasCursor ? 30 : 20; sBgG = hasCursor ? 60 : 50; sBgB = hasCursor ? 30 : 20; }
            else            { sBgR = hasCursor ? 10 : 0;  sBgG = hasCursor ? 55 : 40; sBgB = hasCursor ? 65 : 50; }

            // Background
            SetGCColor(sBgR, sBgG, sBgB, _hasArgbVisual ? _bgAlpha : 255);
            XFillRectangle(_display, _renderDrawable, _gc, spanX, spanY, (uint)spanWidth, (uint)spanHeight);

            // Border
            int borR = hasCursor ? 100 : 60, borG = hasCursor ? 180 : 100, borB = hasCursor ? 200 : 120;
            SetGCColor(borR, borG, borB, 255);
            XDrawRectangle(_display, _renderDrawable, _gc, spanX, spanY, (uint)(spanWidth - 1), (uint)(spanHeight - 1));

            // Cell-ref label
            string cellRef = _grid.GetCellReference(span.anchorRow, span.anchorCol);
            DrawTextWithBg(cellRef, spanX + 2, spanY + 1, 80, 140, 160, sBgR, sBgG, sBgB);

            // Wrapped output content (last N lines that fit)
            if (!string.IsNullOrEmpty(span.content))
            {
                int contentTop = spanY + ch + 2;
                int contentHeight = spanHeight - ch - 4;
                int contentWidth = spanWidth - 8;
                if (contentHeight > 0 && contentWidth > 0)
                {
                    string[] allLines = span.content.Split('\n');
                    int charsPerLine = Math.Max(1, contentWidth / cw);
                    var wrappedLines = new List<string>();
                    foreach (string rawLine in allLines)
                    {
                        if (rawLine.Length <= charsPerLine) wrappedLines.Add(rawLine);
                        else
                            for (int pos = 0; pos < rawLine.Length; pos += charsPerLine)
                                wrappedLines.Add(rawLine.Substring(pos, Math.Min(charsPerLine, rawLine.Length - pos)));
                    }

                    int linesPerWindow = Math.Max(1, contentHeight / ch);
                    int startIdx = Math.Max(0, wrappedLines.Count - linesPerWindow);
                    int count = Math.Min(linesPerWindow, wrappedLines.Count);

                    int contentFgR = span.isCmd ? 100 : 180;
                    int contentFgG = span.isCmd ? 255 : 220;
                    int contentFgB = span.isCmd ? 150 : 255;

                    int lineY = contentTop;
                    for (int li = startIdx; li < startIdx + count && lineY + ch <= contentTop + contentHeight; li++)
                    {
                        string line = wrappedLines[li];
                        if (line.Length > charsPerLine) line = line[..charsPerLine];
                        DrawTextWithBg(line.PadRight(charsPerLine), spanX + 4, lineY,
                            contentFgR, contentFgG, contentFgB, sBgR, sBgG, sBgB);
                        lineY += ch;
                    }
                }
            }
        }

        // Status bar
        string status;
        if (_editMode.IsActive())
        {
            status = _editMode.GetStatusText();
        }
        else if (_searching)
        {
            status = $" Find: {_searchInput}\u2502  (Enter=Search  Esc=Cancel)";
        }
        else
        {
            string cellRef = _grid.GetCellReference(_selectedRow, _selectedCol);
            string value = _grid.GetCellValue(_selectedRow, _selectedCol);
            string valueDisplay = string.IsNullOrEmpty(value) ? "" : $" = {value}";

            string resolvedDisplay = "";
            if (_showResolved && !string.IsNullOrEmpty(value))
            {
                string resolved;
                string? inlineResult = _grid.ResolveInline(_selectedRow, _selectedCol);
                if (inlineResult != null && CellPrefix.IsCommand(inlineResult))
                    resolved = "[running...]";
                else if (inlineResult != null)
                    resolved = CellPrefix.ExpandCellReferences(inlineResult, _grid);
                else
                    resolved = CellPrefix.ExpandCellReferences(value, _grid);
                if (resolved != value)
                    resolvedDisplay = $" \u2192 {resolved}";
            }

            double? sum = _grid.GetColumnSum(_selectedCol);
            string colName = GridManager.GetColumnName(_selectedCol);
            string sumDisplay = sum.HasValue ? $"  \u03a3{colName} = {sum.Value}" : "";
            double? product = _grid.GetRowProduct(_selectedRow);
            string productDisplay = product.HasValue ? $"  \u03a0{(_selectedRow + 1)} = {product.Value}" : "";
            string searchDisplay = _searchTerm != null
                ? $"  \U0001f50d\"{_searchTerm}\" {(_searchMatches.Count > 0 ? $"{_searchMatchIndex + 1}/{_searchMatches.Count}" : "no matches")}"
                : "";
            string f1Label = _showResolved ? "F1: Raw" : "F1: Resolve";
            status = $" {cellRef}{valueDisplay}{resolvedDisplay}{sumDisplay}{productDisplay}{searchDisplay}  |  {f1Label}  F2: Edit  Ctrl+S: Save  Ctrl+Q: Quit";
        }
        int maxChars = _screenWidth / cw;
        status = status.PadRight(maxChars);

        // White background status bar (fully opaque)
        DrawTextWithBg(status, 0, statusY, 0, 0, 0, 255, 255, 255, 255);

        // Blit back buffer to window
        if (_backBuffer != IntPtr.Zero)
            XCopyArea(_display, _backBuffer, _window, _gc, 0, 0, (uint)_screenWidth, (uint)_screenHeight, 0, 0);

        XFlush(_display);
    }

    private int[] GetColumnWidths()
    {
        var widths = new int[_grid.ColumnCount];
        for (int c = 0; c < _grid.ColumnCount; c++)
            widths[c] = _columnWidth;
        return widths;
    }

    /// <summary>
    /// Rebuild the grid for current screen size + column-width setting.
    /// Preserves cell data (skipping desktop-file entries which are re-populated).
    /// </summary>
    private void RebuildGrid()
    {
        // Save unsaved edits before rebuilding so we don't lose them.
        try
        {
            if (_grid.IsDirty)
                _grid.SaveToCsv(_loadedFile ?? AutoSavePath);
        }
        catch { }

        int availableWidth = _screenWidth / _charWidth;
        int availableHeight = _screenHeight / _charHeight - 3;
        var newGrid = new GridManager(availableWidth, availableHeight, _columnWidth);

        for (int r = 0; r < Math.Min(_grid.RowCount, newGrid.RowCount); r++)
            for (int c = 0; c < Math.Min(_grid.ColumnCount, newGrid.ColumnCount); c++)
                if (!_grid.IsFileEntry(r, c))
                    newGrid.SetCellValue(r, c, _grid.GetCellValue(r, c));

        _grid = newGrid;
        _editMode.Grid = _grid;
        _extensionManager.UpdateGrid(_grid);
        _loopManager.UpdateGrid(_grid);

        if (_loadedFile is not null)
            _grid.LoadFromCsv(_loadedFile);
        else if (File.Exists(AutoSavePath))
            _grid.LoadFromCsv(AutoSavePath);

        PopulateDesktopFiles();

        if (_selectedRow >= _grid.RowCount) _selectedRow = _grid.RowCount - 1;
        if (_selectedCol >= _grid.ColumnCount) _selectedCol = _grid.ColumnCount - 1;
        _selection.Clear();
    }

    private void CreateBackBuffer()
    {
        FreeBackBuffer();
        uint depth = _hasArgbVisual ? 32u : (uint)XDefaultDepth(_display, _screen);
        _backBuffer = XCreatePixmap(_display, _window, (uint)_screenWidth, (uint)_screenHeight, depth);
        _backBufferXftDraw = XftDrawCreate(_display, _backBuffer, _visual, _colormap);
    }

    private void FreeBackBuffer()
    {
        if (_backBufferXftDraw != IntPtr.Zero)
        {
            XftDrawDestroy(_backBufferXftDraw);
            _backBufferXftDraw = IntPtr.Zero;
        }
        if (_backBuffer != IntPtr.Zero)
        {
            XFreePixmap(_display, _backBuffer);
            _backBuffer = IntPtr.Zero;
        }
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
        XFillRectangle(_display, _renderDrawable, _gc, x, y, (uint)w, (uint)_charHeight);

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
        XftDrawStringUtf8(_renderXftDraw, ref xftColor, _xftFont, x, y + _fontAscent, utf8, utf8.Length);

        XftColorFree(_display, _visual, _colormap, ref xftColor);
    }

    // ── Input Handling ───────────────────────────────────────────────

    private void HandleKeyPress(ref XKeyEvent keyEvent)
    {
        ulong keysym = XLookupKeysym(ref keyEvent, 0);
        bool ctrl = (keyEvent.state & ControlMask) != 0;
        bool shift = (keyEvent.state & ShiftMask) != 0;

        if (_searching)
        {
            switch (keysym)
            {
                case XK_Return:
                    _searching = false;
                    if (string.IsNullOrEmpty(_searchInput))
                    {
                        _searchTerm = null;
                        _searchMatches.Clear();
                        _searchMatchIndex = -1;
                    }
                    else
                    {
                        _searchTerm = _searchInput;
                        _searchMatches.Clear();
                        for (int r = 0; r < _grid.RowCount; r++)
                            for (int c = 0; c < _grid.ColumnCount; c++)
                                if (_grid.GetCellValue(r, c).Contains(_searchInput, StringComparison.OrdinalIgnoreCase))
                                    _searchMatches.Add((r, c));
                        if (_searchMatches.Count > 0)
                        {
                            _searchMatchIndex = 0;
                            _selectedRow = _searchMatches[0].row;
                            _selectedCol = _searchMatches[0].col;
                        }
                        else
                            _searchMatchIndex = -1;
                    }
                    return;
                case XK_Escape:
                    _searching = false;
                    _searchInput = "";
                    return;
                case XK_BackSpace:
                    if (_searchInput.Length > 0)
                        _searchInput = _searchInput[..^1];
                    return;
                default:
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
                                    _searchInput += ch;
                            }
                        }
                        finally { Marshal.FreeHGlobal(buf); }
                    }
                    return;
            }
        }

        // Edit mode owns its own keys via the IMode dispatch.
        if (_editMode.HandleKeyEventLinux(keysym, ref keyEvent, ctrl))
            return;

        // While search results active, only allow navigation/clear
        if (_searchTerm != null)
        {
            if (keysym == XK_Return && shift && _searchMatches.Count > 0)
            {
                _searchMatchIndex = (_searchMatchIndex - 1 + _searchMatches.Count) % _searchMatches.Count;
                _selectedRow = _searchMatches[_searchMatchIndex].row;
                _selectedCol = _searchMatches[_searchMatchIndex].col;
            }
            else if (keysym == XK_Return && _searchMatches.Count > 0)
            {
                _searchMatchIndex = (_searchMatchIndex + 1) % _searchMatches.Count;
                _selectedRow = _searchMatches[_searchMatchIndex].row;
                _selectedCol = _searchMatches[_searchMatchIndex].col;
            }
            else if (keysym == XK_Escape)
            {
                _searchTerm = null;
                _searchMatches.Clear();
                _searchMatchIndex = -1;
            }
            else if (ctrl && keysym == XK_q)
            {
                _running = false;
            }
            return;
        }

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
                case XK_f:
                    _searching = true;
                    _searchInput = "";
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
            case XK_F1:
                _showResolved = !_showResolved;
                break;
            case XK_F2:
                _grid.SelectCell(_selectedRow, _selectedCol);
                _editMode.Enter();
                break;
            case XK_F3:
                RebuildGrid();
                break;
            case XK_F4:
                if (_columnWidth > MinColumnWidth)
                {
                    _columnWidth -= ColumnWidthStep;
                    RebuildGrid();
                }
                break;
            case XK_F5:
                if (_columnWidth < MaxColumnWidth)
                {
                    _columnWidth += ColumnWidthStep;
                    RebuildGrid();
                }
                break;
            case XK_BackSpace:
                if (_selection.Count > 0)
                {
                    foreach (var (r, c) in _selection) _grid.SetCellValue(r, c, "");
                    _selection.Clear();
                }
                else
                {
                    var val = _grid.GetCellValue(_selectedRow, _selectedCol);
                    if (val.Length > 0) _grid.SetCellValue(_selectedRow, _selectedCol, val[..^1]);
                }
                break;
            case XK_Delete:
                if (_selection.Count > 0)
                {
                    foreach (var (r, c) in _selection) _grid.SetCellValue(r, c, "");
                    _selection.Clear();
                }
                else
                {
                    _grid.SetCellValue(_selectedRow, _selectedCol, "");
                }
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
                // Printable character: auto-enter edit mode and insert (Excel-style).
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
                                _grid.SelectCell(_selectedRow, _selectedCol);
                                _editMode.Enter();
                                _editMode.Insert(ch);
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
                _dragging = true;
                _dragAnchorRow = row;
                _dragAnchorCol = col;
            }

            _selectedRow = row;
            _selectedCol = col;
        }
        else if (buttonEvent.button == 3) // Right click
        {
            OpenAllSelected();
        }
    }

    private void HandleButtonRelease()
    {
        _dragging = false;
    }

    private void HandleMotionNotify(ref XMotionEvent motionEvent)
    {
        if (!_dragging) return;

        int headerRows = 2;
        int row = (motionEvent.y / _charHeight) - headerRows;
        row = Math.Clamp(row, 0, _grid.RowCount - 1);

        int[] colWidths = GetColumnWidths();
        int x = RowHeaderWidth * _charWidth;
        int col = _grid.ColumnCount - 1; // default to last col if past end
        for (int c = 0; c < _grid.ColumnCount; c++)
        {
            int colPx = colWidths[c] * _charWidth;
            if (motionEvent.x < x + colPx) { col = c; break; }
            x += colPx;
        }
        col = Math.Clamp(col, 0, _grid.ColumnCount - 1);

        _selectedRow = row;
        _selectedCol = col;
        SelectRectangle(_dragAnchorRow, _dragAnchorCol, row, col);
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
        string filename = _loadedFile ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "spreadsheet.csv");
        _grid.SaveToCsv(filename);
        _loadedFile = filename;
    }

    // ── Dispose ──────────────────────────────────────────────────────

    public void Dispose()
    {
        _loopManager.Dispose();
        _autoSaveTimer?.Dispose();

        // Autosave on exit
        try
        {
            Directory.CreateDirectory(StateDir);
            _grid.SaveToCsv(AutoSavePath);
        }
        catch { }

        _inlineProcesses.Dispose();
        _extensionManager.Dispose();

        FreeBackBuffer();

        if (_xftDraw != IntPtr.Zero)
            XftDrawDestroy(_xftDraw);
        if (_xftFont != IntPtr.Zero)
            XftFontClose(_display, _xftFont);
        if (_gc != IntPtr.Zero)
            XFreeGC(_display, _gc);
        if (_window != IntPtr.Zero)
            XDestroyWindow(_display, _window);
    }

    private void OpenAllSelected()
    {
        var cells = new HashSet<(int row, int col)>(_selection)
        {
            (_selectedRow, _selectedCol)
        };

        bool opened = false;
        foreach (var (r, c) in cells)
            opened |= ActivateCellAt(r, c);

        if (!opened && _selectedRow < _grid.RowCount - 1)
            _selectedRow++;
    }

    /// <summary>
    /// Activates a single cell as if the user pressed Enter on it.
    /// Handles extension reactivation, inline command reruns, file/command/hyperlink opening.
    /// Returns true if the cell was meaningfully activated.
    /// </summary>
    private bool ActivateCellAt(int row, int col)
    {
        if (row < 0 || row >= _grid.RowCount || col < 0 || col >= _grid.ColumnCount)
            return false;

        _extensionManager.ReactivateCell(row, col);

        if (TryRerunInlineCommand(row, col))
            return true;

        if (_grid.IsFileEntry(row, col))
        {
            string path = _grid.GetFilePath(row, col);
            if (!string.IsNullOrEmpty(path))
                try
                {
                    var psi = new ProcessStartInfo("xdg-open") { UseShellExecute = false };
                    psi.ArgumentList.Add(path);
                    Process.Start(psi);
                    return true;
                } catch { }
        }
        else
        {
            string val = _grid.GetCellValue(row, col);
            if (IsCommand(val))
            {
                RunCommand(val);
                return true;
            }
            else if (IsHyperlink(val))
                try
                {
                    var psi = new ProcessStartInfo("xdg-open") { UseShellExecute = false };
                    psi.ArgumentList.Add(val);
                    Process.Start(psi);
                    return true;
                } catch { }
        }

        return false;
    }

    /// <summary>
    /// If the cell is an inline ref pointing to a command, kill the cached process
    /// so the next render restarts it. Returns true if it handled the action.
    /// </summary>
    private bool TryRerunInlineCommand(int row, int col)
    {
        string val = _grid.GetCellValue(row, col);
        if (!CellPrefix.IsInline(val)) return false;
        string? resolved = _grid.ResolveInline(row, col);
        if (resolved == null || !CellPrefix.IsCommand(resolved)) return false;
        _inlineProcesses.StopProcess(row, col);
        return true;
    }
}
