namespace ExcelConsole.Terminal;

/// <summary>
/// Terminal screen model fed with raw PTY output bytes. All members thread-safe
/// (PTY reader thread writes; UI thread reads).
/// </summary>
internal interface IVtScreen : IDisposable
{
    /// <summary>Feed raw bytes exactly as read from the PTY (may split UTF-8 / escape sequences across calls).</summary>
    void Write(ReadOnlySpan<byte> data);

    /// <summary>Resize the emulated grid (cells). Values &lt; 1 are clamped to 1, &gt; 65535 clamped to 65535.</summary>
    void Resize(int cols, int rows);

    /// <summary>True if Write/Resize happened since the last GetText() call.</summary>
    bool IsDirty { get; }

    /// <summary>
    /// Plain text of scrollback + visible screen, lines joined with '\n', no '\r',
    /// trailing whitespace per line trimmed, trailing empty lines removed. "" when empty.
    /// Clears IsDirty. Cheap when not dirty (returns cached string).
    /// </summary>
    string GetText();
}
