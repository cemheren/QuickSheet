namespace ExcelConsole.Terminal;

/// <summary>
/// One interactive terminal living in a block of cells (a t: cell). Owns the shell's pty and
/// the libghostty terminal that interprets its output. Platform-neutral: hosts translate native
/// key events to <see cref="TermKey"/> and draw <see cref="TerminalFrame"/>s.
/// </summary>
internal sealed class TerminalPane : IDisposable
{
    private readonly GhosttyTerminal _term;
    private readonly ITerminalPty? _pty;
    private bool _exitSeen;

    public string? Error { get; }
    public bool HasExited => _pty?.HasExited ?? true;
    public int ExitCode => _pty?.ExitCode ?? -1;
    public string Title => _term.Title;
    public int Cols => _term.Cols;
    public int Rows => _term.Rows;

    private TerminalPane(GhosttyTerminal term, ITerminalPty? pty, string? error)
    {
        _term = term;
        _pty = pty;
        Error = error;
    }

    /// <summary>
    /// Creates and starts a pty running the user's shell at cols×rows, delivering output to the
    /// callback (reader thread, reused buffer). Returns null on failure. Supplied by the platform host.
    /// </summary>
    public delegate ITerminalPty? PtyFactory(int cols, int rows, Action<byte[], int> onOutput);

    /// <summary>Start the user's shell in a cols×rows terminal. Null if libghostty-vt is unavailable.</summary>
    public static TerminalPane? Start(PtyFactory ptyFactory, int cols, int rows, int cellWidthPx, int cellHeightPx,
        int defaultFg, int defaultBg)
    {
        if (!GhosttyTerminal.IsSupported) return null;
        GhosttyTerminal term;
        try { term = new GhosttyTerminal(cols, rows, cellWidthPx, cellHeightPx, defaultFg, defaultBg); }
        catch (Exception ex)
        {
            try { Console.Error.WriteLine($"[QuickSheet] terminal pane: {ex.Message}"); } catch { }
            return null;
        }
        ITerminalPty? pty = null;
        // Replies the terminal generates while parsing output go back to the program.
        term.SetPtyWriter(bytes => pty?.Write(bytes));
        try { pty = ptyFactory(term.Cols, term.Rows, (buf, n) => term.Feed(buf.AsSpan(0, n))); }
        catch { pty = null; }
        if (pty == null)
        {
            term.Feed("\x1b[31mFailed to start shell.\x1b[0m\r\n"u8);
            return new TerminalPane(term, null, "failed to start shell");
        }
        return new TerminalPane(term, pty, null);
    }

    /// <summary>True when there is something new to draw (output, resize, or the shell just exited).</summary>
    public bool NeedsRedraw
    {
        get
        {
            if (_term.IsDirty) return true;
            if (!_exitSeen && HasExited) { _exitSeen = true; return true; }
            return false;
        }
    }

    public TerminalFrame? Snapshot() => _term.Snapshot();

    public void Resize(int cols, int rows, int cellWidthPx, int cellHeightPx)
    {
        if (_term.Resize(cols, rows, cellWidthPx, cellHeightPx))
            _pty?.Resize(_term.Cols, _term.Rows);
    }

    /// <summary>Send a key press. Returns true if it produced input.</summary>
    public bool SendKey(TermKey key, TermMods mods, string? text, uint unshiftedCodepoint,
        TermMods consumedMods = TermMods.None, bool repeat = false)
    {
        if (_pty == null || HasExited) return false;
        byte[]? bytes = _term.EncodeKey(key, mods, text, unshiftedCodepoint, consumedMods, repeat);
        if (bytes == null || bytes.Length == 0) return false;
        _term.ScrollToBottom();
        _pty.Write(bytes);
        return true;
    }

    /// <summary>Send text the user typed that no key maps to (e.g. IME / dead-key composition).</summary>
    public void SendText(string text)
    {
        if (_pty == null || HasExited || string.IsNullOrEmpty(text)) return;
        _term.ScrollToBottom();
        _pty.Write(System.Text.Encoding.UTF8.GetBytes(text));
    }

    public void Paste(string text)
    {
        if (_pty == null || HasExited) return;
        byte[]? bytes = _term.EncodePaste(text);
        if (bytes == null) return;
        _term.ScrollToBottom();
        _pty.Write(bytes);
    }

    public void Scroll(int deltaRows) => _term.ScrollViewport(deltaRows);

    public void Dispose()
    {
        _pty?.Dispose();
        _term.Dispose();
    }
}
