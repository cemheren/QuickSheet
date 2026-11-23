using System.Runtime.InteropServices;
using System.Text;
using static ExcelConsole.Terminal.GhosttyNative;

namespace ExcelConsole.Terminal;

/// <summary>
/// An interactive libghostty-vt terminal for a t: pane. Unlike <see cref="GhosttyVtScreen"/>
/// (plain-text output of i: cells), this one answers the program's queries through the pty,
/// encodes keys and pastes, and exposes the styled screen as a <see cref="TerminalFrame"/>.
///
/// Thread-safety: the pty reader thread calls <see cref="Feed"/>; the UI thread calls everything
/// else. All native calls are serialized by one lock.
/// </summary>
internal sealed class GhosttyTerminal : IDisposable
{
    private readonly object _lock = new();
    private IntPtr _terminal;
    private IntPtr _renderState;
    private IntPtr _rowIterator;
    private IntPtr _rowCells;
    private IntPtr _keyEncoder;
    private IntPtr _keyEvent;
    private IntPtr _scratch;      // small reusable native block for out-params
    private IntPtr _textBuf;      // grapheme UTF-8 output buffer
    private IntPtr _keyText;      // key event UTF-8 input buffer (must outlive the encode call)
    private const int TextBufSize = 64;
    private const int KeyTextSize = 64;

    // Keep the callback delegates alive as long as the terminal can call them.
    private readonly WritePtyFn _writePtyFn;
    private readonly TitleChangedFn _titleChangedFn;

    private Action<byte[]>? _ptyWriter;
    private volatile bool _dirty = true;
    private TerminalFrame? _frame;
    private int _cols, _rows;
    private bool _disposed;

    private static readonly string[] s_ascii = BuildAsciiStrings();

    public string Title { get; private set; } = "";
    public int Cols { get { lock (_lock) return _cols; } }
    public int Rows { get { lock (_lock) return _rows; } }

    /// <summary>True when the screen changed since the last <see cref="Snapshot"/>.</summary>
    public bool IsDirty => _dirty;

    public static bool IsSupported => GhosttyNative.IsAvailable;

    public GhosttyTerminal(int cols, int rows, int cellWidthPx, int cellHeightPx, int defaultFg, int defaultBg,
        int maxScrollbackLines = 5000)
    {
        if (!GhosttyNative.IsAvailable)
            throw new InvalidOperationException($"libghostty-vt unavailable: {GhosttyNative.LoadError}");

        _cols = Clamp(cols);
        _rows = Clamp(rows);
        _writePtyFn = OnWritePty;
        _titleChangedFn = OnTitleChanged;

        try
        {
            Check(ghostty_terminal_new(IntPtr.Zero, out _terminal, (ushort)_cols, (ushort)_rows), "ghostty_terminal_new");
            Check(TerminalSetNuint(_terminal, OptScrollbackMaxLines, (nuint)maxScrollbackLines), "set SCROLLBACK_MAX_LINES");
            Check(TerminalSetNuint(_terminal, OptScrollbackMaxBytes, (nuint)Math.Clamp(
                (long)_cols * 16 * (maxScrollbackLines + _rows), 8L << 20, 256L << 20)), "set SCROLLBACK_MAX_BYTES");
            Check(ghostty_terminal_set(_terminal, OptWritePty, Marshal.GetFunctionPointerForDelegate(_writePtyFn)), "set WRITE_PTY");
            Check(ghostty_terminal_set(_terminal, OptTitleChanged, Marshal.GetFunctionPointerForDelegate(_titleChangedFn)), "set TITLE_CHANGED");
            SetColor(OptColorForeground, defaultFg);
            SetColor(OptColorBackground, defaultBg);
            if (cellWidthPx > 0 && cellHeightPx > 0)
                ghostty_terminal_resize(_terminal, (ushort)_cols, (ushort)_rows, (uint)cellWidthPx, (uint)cellHeightPx);

            Check(ghostty_render_state_new(IntPtr.Zero, out _renderState), "ghostty_render_state_new");
            Check(ghostty_render_state_row_iterator_new(IntPtr.Zero, out _rowIterator), "row_iterator_new");
            Check(ghostty_render_state_row_cells_new(IntPtr.Zero, out _rowCells), "row_cells_new");
            Check(ghostty_key_encoder_new(IntPtr.Zero, out _keyEncoder), "ghostty_key_encoder_new");
            Check(ghostty_key_event_new(IntPtr.Zero, out _keyEvent), "ghostty_key_event_new");

            _scratch = Marshal.AllocHGlobal(256);
            _textBuf = Marshal.AllocHGlobal(TextBufSize);
            _keyText = Marshal.AllocHGlobal(KeyTextSize);
        }
        catch
        {
            FreeHandles();
            throw;
        }
    }

    /// <summary>Where replies from the terminal (device reports, mode queries) and encoded input go.</summary>
    public void SetPtyWriter(Action<byte[]> writer) => _ptyWriter = writer;

    // ── Output from the program ─────────────────────────────────────

    /// <summary>Feed raw pty output. Called from the pty reader thread.</summary>
    public void Feed(ReadOnlySpan<byte> data)
    {
        if (data.IsEmpty) return;
        lock (_lock)
        {
            if (_disposed) return;
            ghostty_terminal_vt_write(_terminal, ref MemoryMarshal.GetReference(data), (nuint)data.Length);
            _dirty = true;
        }
    }

    private void OnWritePty(IntPtr terminal, IntPtr userdata, IntPtr data, nuint len)
    {
        if (data == IntPtr.Zero || len == 0) return;
        var bytes = new byte[(int)len];
        Marshal.Copy(data, bytes, 0, bytes.Length);
        try { _ptyWriter?.Invoke(bytes); } catch { }
    }

    private void OnTitleChanged(IntPtr terminal, IntPtr userdata)
    {
        // Runs inside Feed (lock held). The returned string is borrowed: copy it now.
        IntPtr p = _scratch;
        Marshal.StructureToPtr(new GString(), p, false);
        if (ghostty_terminal_get(_terminal, 12 /* DATA_TITLE */, p) != Success) return;
        var s = Marshal.PtrToStructure<GString>(p);
        Title = s.Ptr == IntPtr.Zero || s.Len == 0 ? "" : Marshal.PtrToStringUTF8(s.Ptr, (int)s.Len);
    }

    // ── Size ─────────────────────────────────────────────────────────

    /// <summary>Resize the emulated screen. Returns true if the size changed.</summary>
    public bool Resize(int cols, int rows, int cellWidthPx, int cellHeightPx)
    {
        cols = Clamp(cols);
        rows = Clamp(rows);
        lock (_lock)
        {
            if (_disposed || (cols == _cols && rows == _rows)) return false;
            if (ghostty_terminal_resize(_terminal, (ushort)cols, (ushort)rows,
                    (uint)Math.Max(1, cellWidthPx), (uint)Math.Max(1, cellHeightPx)) != Success)
                return false;
            _cols = cols;
            _rows = rows;
            _dirty = true;
            return true;
        }
    }

    // ── Input to the program ─────────────────────────────────────────

    /// <summary>
    /// Encode a key press for the program, honoring its current modes (application cursor keys,
    /// Kitty keyboard protocol, ...). <paramref name="text"/> is the unmodified text the key types
    /// with the current layout (no control characters); empty for function keys.
    /// Returns null when the key produces no input.
    /// </summary>
    public byte[]? EncodeKey(TermKey key, TermMods mods, string? text, uint unshiftedCodepoint,
        TermMods consumedMods = TermMods.None, bool repeat = false)
    {
        lock (_lock)
        {
            if (_disposed) return null;
            ghostty_key_encoder_set_from_terminal(_keyEncoder, _terminal);
            ghostty_key_event_set_action(_keyEvent, repeat ? KeyActionRepeat : KeyActionPress);
            ghostty_key_event_set_key(_keyEvent, (int)key);
            ghostty_key_event_set_mods(_keyEvent, (ushort)mods);
            ghostty_key_event_set_consumed_mods(_keyEvent, (ushort)consumedMods);
            ghostty_key_event_set_unshifted_codepoint(_keyEvent, unshiftedCodepoint);

            int textLen = 0;
            if (!string.IsNullOrEmpty(text))
            {
                byte[] utf8 = Encoding.UTF8.GetBytes(text);
                textLen = Math.Min(utf8.Length, KeyTextSize);
                Marshal.Copy(utf8, 0, _keyText, textLen);
            }
            ghostty_key_event_set_utf8(_keyEvent, new GString { Ptr = _keyText, Len = (nuint)textLen });

            byte[] buf = new byte[128];
            int r = ghostty_key_encoder_encode(_keyEncoder, _keyEvent, buf, (nuint)buf.Length, out nuint written);
            if (r == OutOfSpace)
            {
                buf = new byte[(int)written];
                r = ghostty_key_encoder_encode(_keyEncoder, _keyEvent, buf, (nuint)buf.Length, out written);
            }
            ghostty_key_event_set_utf8(_keyEvent, new GString());
            if (r != Success || written == 0) return null;
            return buf.AsSpan(0, (int)written).ToArray();
        }
    }

    /// <summary>Encode pasted text (bracketed when the program asked for it; unsafe control bytes stripped).</summary>
    public byte[]? EncodePaste(string text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        lock (_lock)
        {
            if (_disposed) return null;
            bool bracketed = false;
            Marshal.StructureToPtr(new ModeConfig { Mode = ModeBracketedPaste }, _scratch, false);
            if (ghostty_terminal_get(_terminal, DataMode, _scratch) == Success)
                bracketed = Marshal.PtrToStructure<ModeConfig>(_scratch).Value;

            byte[] data = Encoding.UTF8.GetBytes(text);
            int r = ghostty_paste_encode(data, (nuint)data.Length, bracketed, null, 0, out nuint need);
            if (r != OutOfSpace && r != Success) return null;
            if (need == 0) return null;
            byte[] buf = new byte[(int)need];
            // The encoder rewrites `data` in place, so start again from a fresh copy.
            data = Encoding.UTF8.GetBytes(text);
            r = ghostty_paste_encode(data, (nuint)data.Length, bracketed, buf, (nuint)buf.Length, out nuint written);
            return r == Success ? buf.AsSpan(0, (int)written).ToArray() : null;
        }
    }

    /// <summary>Scroll the viewport through scrollback. Negative = up (older output).</summary>
    public void ScrollViewport(int deltaRows)
    {
        lock (_lock)
        {
            if (_disposed) return;
            ghostty_terminal_scroll_viewport(_terminal, new ScrollViewport { Tag = ScrollViewportDelta, Delta = deltaRows });
            _dirty = true;
        }
    }

    public void ScrollToBottom()
    {
        lock (_lock)
        {
            if (_disposed) return;
            ghostty_terminal_scroll_viewport(_terminal, new ScrollViewport { Tag = ScrollViewportBottom });
            _dirty = true;
        }
    }

    // ── Rendering ────────────────────────────────────────────────────

    /// <summary>
    /// Snapshot of the visible screen. Rebuilt only when the terminal changed; otherwise the
    /// previous frame is returned. Null only if the very first snapshot failed.
    /// </summary>
    public TerminalFrame? Snapshot()
    {
        lock (_lock)
        {
            if (_disposed) return _frame;
            if (!_dirty && _frame != null) return _frame;
            _dirty = false;

            if (ghostty_render_state_update(_renderState, _terminal) != Success) return _frame;

            int cols = GetU16(RsDataCols, _cols);
            int rows = GetU16(RsDataRows, _rows);
            int defFg = GetColor(_renderState, RsDataColorForeground) ?? 0xd0d0d0;
            int defBg = GetColor(_renderState, RsDataColorBackground) ?? 0x000000;

            bool cursorVisible = GetBool(RsDataCursorVisible) && GetBool(RsDataCursorViewportHasValue);
            int cx = cursorVisible ? GetU16(RsDataCursorViewportX, 0) : 0;
            int cy = cursorVisible ? GetU16(RsDataCursorViewportY, 0) : 0;

            var cells = new FrameCell[cols * rows];
            for (int i = 0; i < cells.Length; i++) { cells[i].Fg = defFg; cells[i].Bg = defBg; }

            IntPtr slot = _scratch;                    // handle out-param
            IntPtr bufStruct = _scratch + 16;          // GhosttyBuffer (24 bytes)
            IntPtr styleOut = _scratch + 64;           // GhosttyStyle (16 bytes)
            IntPtr colorOut = _scratch + 96;           // GhosttyColorRgb (3 bytes)
            IntPtr boolOut = _scratch + 104;           // bool

            Marshal.WriteIntPtr(slot, _rowIterator);
            if (ghostty_render_state_get(_renderState, RsDataRowIterator, slot) == Success)
            {
                int y = 0;
                while (y < rows && ghostty_render_state_row_iterator_next(_rowIterator))
                {
                    Marshal.WriteIntPtr(slot, _rowCells);
                    if (ghostty_render_state_row_get(_rowIterator, RsRowDataCells, slot) != Success) { y++; continue; }

                    int x = 0;
                    while (x < cols && ghostty_render_state_row_cells_next(_rowCells))
                    {
                        ref FrameCell cell = ref cells[y * cols + x];

                        // Text
                        Marshal.StructureToPtr(new GBuffer { Ptr = _textBuf, Cap = TextBufSize, Len = 0 }, bufStruct, false);
                        if (ghostty_render_state_row_cells_get(_rowCells, RsCellGraphemesUtf8, bufStruct) == Success)
                        {
                            int len = (int)Marshal.PtrToStructure<GBuffer>(bufStruct).Len;
                            if (len == 1)
                            {
                                byte b = Marshal.ReadByte(_textBuf);
                                cell.Text = b < 128 ? s_ascii[b] : null;
                            }
                            else if (len > 1)
                                cell.Text = Marshal.PtrToStringUTF8(_textBuf, len);
                        }

                        // Colors: background can be set without styling (erase with color).
                        int bg = defBg, fg = defFg;
                        if (ghostty_render_state_row_cells_get(_rowCells, RsCellBgColor, colorOut) == Success)
                            bg = ReadRgb(colorOut);

                        Marshal.WriteByte(boolOut, 0);
                        ghostty_render_state_row_cells_get(_rowCells, RsCellHasStyling, boolOut);
                        if (Marshal.ReadByte(boolOut) != 0)
                        {
                            if (ghostty_render_state_row_cells_get(_rowCells, RsCellFgColor, colorOut) == Success)
                                fg = ReadRgb(colorOut);
                            if (ghostty_render_state_row_cells_get(_rowCells, RsCellStyle, styleOut) == Success)
                            {
                                ushort flags = Marshal.PtrToStructure<Style>(styleOut).Flags;
                                CellAttr a = CellAttr.None;
                                if ((flags & StyleBold) != 0) a |= CellAttr.Bold;
                                if ((flags & StyleItalic) != 0) a |= CellAttr.Italic;
                                if ((flags & StyleFaint) != 0) a |= CellAttr.Faint;
                                if ((flags & StyleStrikethrough) != 0) a |= CellAttr.Strike;
                                if ((flags & StyleInvisible) != 0) a |= CellAttr.Invisible;
                                if ((flags & StyleUnderlineMask) != 0) a |= CellAttr.Underline;
                                if ((flags & StyleInverse) != 0) (fg, bg) = (bg, fg);
                                if ((a & CellAttr.Faint) != 0) fg = Blend(fg, bg);
                                cell.Attr = a;
                            }
                        }
                        cell.Fg = fg;
                        cell.Bg = bg;
                        x++;
                    }
                    y++;
                }
            }

            _frame = new TerminalFrame
            {
                Cols = cols,
                Rows = rows,
                Cells = cells,
                DefaultFg = defFg,
                DefaultBg = defBg,
                CursorVisible = cursorVisible && cx < cols && cy < rows,
                CursorX = cx,
                CursorY = cy,
            };
            return _frame;
        }
    }

    // ── helpers ──────────────────────────────────────────────────────

    private int GetU16(int key, int fallback)
    {
        Marshal.WriteInt16(_scratch + 128, 0);
        return ghostty_render_state_get(_renderState, key, _scratch + 128) == Success
            ? (ushort)Marshal.ReadInt16(_scratch + 128) : fallback;
    }

    private bool GetBool(int key)
    {
        Marshal.WriteByte(_scratch + 136, 0);
        return ghostty_render_state_get(_renderState, key, _scratch + 136) == Success
            && Marshal.ReadByte(_scratch + 136) != 0;
    }

    private int? GetColor(IntPtr state, int key)
    {
        IntPtr p = _scratch + 144;
        return ghostty_render_state_get(state, key, p) == Success ? ReadRgb(p) : null;
    }

    private static int ReadRgb(IntPtr p) =>
        (Marshal.ReadByte(p) << 16) | (Marshal.ReadByte(p, 1) << 8) | Marshal.ReadByte(p, 2);

    private static int Blend(int fg, int bg)
    {
        int r = (((fg >> 16) & 0xff) + ((bg >> 16) & 0xff)) / 2;
        int g = (((fg >> 8) & 0xff) + ((bg >> 8) & 0xff)) / 2;
        int b = ((fg & 0xff) + (bg & 0xff)) / 2;
        return (r << 16) | (g << 8) | b;
    }

    private void SetColor(int option, int rgb)
    {
        IntPtr p = Marshal.AllocHGlobal(4);
        try
        {
            Marshal.WriteByte(p, 0, (byte)((rgb >> 16) & 0xff));
            Marshal.WriteByte(p, 1, (byte)((rgb >> 8) & 0xff));
            Marshal.WriteByte(p, 2, (byte)(rgb & 0xff));
            ghostty_terminal_set(_terminal, option, p);
        }
        finally { Marshal.FreeHGlobal(p); }
    }

    private static string[] BuildAsciiStrings()
    {
        var a = new string[128];
        for (int i = 0; i < 128; i++) a[i] = i < 32 || i == 127 ? " " : ((char)i).ToString();
        return a;
    }

    private static int Clamp(int v) => v < 1 ? 1 : v > 4096 ? 4096 : v;

    private static void Check(int result, string what)
    {
        if (result != Success) throw new InvalidOperationException($"{what} failed: {result}");
    }

    private void FreeHandles()
    {
        if (_keyEvent != IntPtr.Zero) { ghostty_key_event_free(_keyEvent); _keyEvent = IntPtr.Zero; }
        if (_keyEncoder != IntPtr.Zero) { ghostty_key_encoder_free(_keyEncoder); _keyEncoder = IntPtr.Zero; }
        if (_rowCells != IntPtr.Zero) { ghostty_render_state_row_cells_free(_rowCells); _rowCells = IntPtr.Zero; }
        if (_rowIterator != IntPtr.Zero) { ghostty_render_state_row_iterator_free(_rowIterator); _rowIterator = IntPtr.Zero; }
        if (_renderState != IntPtr.Zero) { ghostty_render_state_free(_renderState); _renderState = IntPtr.Zero; }
        if (_terminal != IntPtr.Zero) { ghostty_terminal_free(_terminal); _terminal = IntPtr.Zero; }
        if (_scratch != IntPtr.Zero) { Marshal.FreeHGlobal(_scratch); _scratch = IntPtr.Zero; }
        if (_textBuf != IntPtr.Zero) { Marshal.FreeHGlobal(_textBuf); _textBuf = IntPtr.Zero; }
        if (_keyText != IntPtr.Zero) { Marshal.FreeHGlobal(_keyText); _keyText = IntPtr.Zero; }
    }

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;
            _ptyWriter = null;
            FreeHandles();
        }
    }
}
