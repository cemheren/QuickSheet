using System.Runtime.InteropServices;

namespace ExcelConsole.Terminal;

/// <summary>IVtScreen backed by libghostty-vt.</summary>
internal sealed class GhosttyVtScreen : IVtScreen
{
    private readonly object _lock = new();
    private IntPtr _terminal;
    private IntPtr _formatter;
    private int _cols, _rows;
    private bool _dirty = true;
    private string? _cache;
    private bool _disposed;
    private readonly int _maxScrollbackLines;
    private int _formatFailures;

    public GhosttyVtScreen(int cols, int rows, int maxScrollbackLines)
    {
        _cols = Clamp(cols);
        _rows = Clamp(rows);
        _maxScrollbackLines = Math.Max(0, maxScrollbackLines);
        int r = GhosttyNative.ghostty_terminal_new(IntPtr.Zero, out _terminal, (ushort)_cols, (ushort)_rows);
        if (r != GhosttyNative.Success || _terminal == IntPtr.Zero)
        {
            _terminal = IntPtr.Zero;
            throw new InvalidOperationException($"ghostty_terminal_new failed: {r}");
        }
        try
        {
            Check(GhosttyNative.TerminalSetNuint(_terminal, GhosttyNative.OptScrollbackMaxLines,
                (nuint)Math.Max(0, maxScrollbackLines)), "set SCROLLBACK_MAX_LINES");
            Check(GhosttyNative.TerminalSetNuint(_terminal, GhosttyNative.OptScrollbackMaxBytes,
                ScrollbackBytes(_cols, _rows, _maxScrollbackLines)), "set SCROLLBACK_MAX_BYTES");
#if PLATFORM_WINDOWS
            // ConPTY keeps its own screen buffer without scrollback; don't pull scrollback on resize.
            // Tuning option only: ignore failure.
            _ = GhosttyNative.TerminalSetBool(_terminal, GhosttyNative.OptResizePullScrollback, false);
#endif
            Check(GhosttyNative.ghostty_formatter_terminal_new(IntPtr.Zero, out _formatter, _terminal),
                "ghostty_formatter_terminal_new");
            Check(GhosttyNative.FormatterSetInt(_formatter, GhosttyNative.FormatterOptEmit, GhosttyNative.FormatPlain), "set EMIT");
            Check(GhosttyNative.FormatterSetBool(_formatter, GhosttyNative.FormatterOptUnwrap, true), "set UNWRAP");
            Check(GhosttyNative.FormatterSetBool(_formatter, GhosttyNative.FormatterOptTrim, true), "set TRIM");
        }
        catch
        {
            FreeHandles();
            throw;
        }
    }

    // A ghostty cell is 8 bytes and rows are full width, so a fixed byte cap would truncate wide
    // (e.g. 4096-col JSON) terminals well below the line target. Scale the cap with the size;
    // x2 leaves slack for style/grapheme metadata.
    private static nuint ScrollbackBytes(int cols, int rows, int maxLines)
    {
        ulong bytes = (ulong)cols * 8 * 2 * (ulong)(maxLines + rows);
        const ulong min = 8UL * 1024 * 1024, max = 512UL * 1024 * 1024;
        return (nuint)Math.Clamp(bytes, min, max);
    }

    private static int Clamp(int v) => v < 1 ? 1 : v > 65535 ? 65535 : v;

    private static void Check(int result, string what)
    {
        if (result != GhosttyNative.Success)
            throw new InvalidOperationException($"{what} failed: {result}");
    }

    private void FreeHandles()
    {
        if (_formatter != IntPtr.Zero) { GhosttyNative.ghostty_formatter_free(_formatter); _formatter = IntPtr.Zero; }
        if (_terminal != IntPtr.Zero) { GhosttyNative.ghostty_terminal_free(_terminal); _terminal = IntPtr.Zero; }
    }

    public bool IsDirty { get { lock (_lock) return _dirty; } }

    public void Write(ReadOnlySpan<byte> data)
    {
        if (data.IsEmpty) return;
        lock (_lock)
        {
            if (_disposed) return;
            GhosttyNative.ghostty_terminal_vt_write(_terminal, ref MemoryMarshal.GetReference(data), (nuint)data.Length);
            _dirty = true;
        }
    }

    public void Resize(int cols, int rows)
    {
        cols = Clamp(cols);
        rows = Clamp(rows);
        lock (_lock)
        {
            if (_disposed || (cols == _cols && rows == _rows)) return;
            if (GhosttyNative.ghostty_terminal_resize(_terminal, (ushort)cols, (ushort)rows, 8, 16) == GhosttyNative.Success)
            {
                if (cols != _cols)
                    GhosttyNative.TerminalSetNuint(_terminal, GhosttyNative.OptScrollbackMaxBytes,
                        ScrollbackBytes(cols, rows, _maxScrollbackLines));
                _cols = cols;
                _rows = rows;
            }
            _dirty = true;
        }
    }

    public string GetText()
    {
        lock (_lock)
        {
            if (_disposed) return _cache ?? "";
            if (!_dirty && _cache != null) return _cache;

            string text = "";
            int r = GhosttyNative.ghostty_formatter_format_alloc(_formatter, IntPtr.Zero, out IntPtr ptr, out nuint len);
            if (r == GhosttyNative.Success && ptr != IntPtr.Zero)
            {
                try
                {
                    if (len > 0)
                        text = Marshal.PtrToStringUTF8(ptr, checked((int)len));
                }
                finally { GhosttyNative.ghostty_free(IntPtr.Zero, ptr, len); }
            }
            else if (r != GhosttyNative.Success)
            {
                // keep previous cache on failure; retry while dirty, but give up after 3 in a row
                // so the render loop can idle
                if (++_formatFailures >= 3) _dirty = false;
                return _cache ?? "";
            }

            _formatFailures = 0;
            // The formatter emits leading blank rows; single-cell renderers show up to the first '\n'.
            text = text.Replace("\r", "").TrimStart('\n').TrimEnd('\n', ' ');
            _cache = text;
            _dirty = false;
            return text;
        }
    }

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;
            FreeHandles();
        }
    }
}
