namespace ExcelConsole.Terminal;

[Flags]
internal enum CellAttr : byte
{
    None = 0,
    Bold = 1 << 0,
    Italic = 1 << 1,
    Faint = 1 << 2,
    Underline = 1 << 3,
    Strike = 1 << 4,
    Invisible = 1 << 5,
}

/// <summary>One rendered terminal cell. Colors are 0xRRGGBB, already resolved (palette, inverse).</summary>
internal struct FrameCell
{
    /// <summary>Grapheme text; null for an empty cell or the spacer half of a wide character.</summary>
    public string? Text;
    public int Fg;
    public int Bg;
    public CellAttr Attr;
}

/// <summary>
/// Immutable snapshot of the visible terminal screen, built from libghostty's render state.
/// Platform renderers draw it; nothing here is platform-specific.
/// </summary>
internal sealed class TerminalFrame
{
    public int Cols { get; init; }
    public int Rows { get; init; }
    public FrameCell[] Cells { get; init; } = Array.Empty<FrameCell>();
    public int DefaultFg { get; init; }
    public int DefaultBg { get; init; }
    public bool CursorVisible { get; init; }
    public int CursorX { get; init; }
    public int CursorY { get; init; }

    public ref FrameCell this[int row, int col] => ref Cells[row * Cols + col];

    public static (int r, int g, int b) Rgb(int c) => ((c >> 16) & 0xff, (c >> 8) & 0xff, c & 0xff);
}
