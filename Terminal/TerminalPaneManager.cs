namespace ExcelConsole.Terminal;

/// <summary>
/// Keeps one <see cref="TerminalPane"/> per t: anchor cell. Hosts call <see cref="Sync"/> each
/// render with the anchors currently on the grid; panes whose cell was cleared are closed.
/// Panes are not persisted: after a QuickSheet restart each t: cell starts a fresh shell.
/// UI-thread only.
/// </summary>
internal sealed class TerminalPaneManager : IDisposable
{
    private readonly Dictionary<(int row, int col), TerminalPane> _panes = new();
    private readonly TerminalPane.PtyFactory _ptyFactory;

    public TerminalPaneManager(TerminalPane.PtyFactory ptyFactory) => _ptyFactory = ptyFactory;

    public int Count => _panes.Count;

    public TerminalPane? Get(int row, int col) => _panes.TryGetValue((row, col), out var p) ? p : null;

    /// <summary>
    /// Returns the pane for an anchor, starting its shell on first use, and keeps its size in
    /// sync with the cell block. Null when interactive terminals are unsupported (no libghostty-vt).
    /// </summary>
    public TerminalPane? Ensure(int row, int col, int cols, int rows, int cellWidthPx, int cellHeightPx,
        int defaultFg, int defaultBg)
    {
        if (_panes.TryGetValue((row, col), out var pane))
        {
            pane.Resize(cols, rows, cellWidthPx, cellHeightPx);
            return pane;
        }
        pane = TerminalPane.Start(_ptyFactory, cols, rows, cellWidthPx, cellHeightPx, defaultFg, defaultBg);
        if (pane != null) _panes[(row, col)] = pane;
        return pane;
    }

    /// <summary>Close the pane so the next <see cref="Ensure"/> starts a fresh shell.</summary>
    public void Restart(int row, int col)
    {
        if (_panes.Remove((row, col), out var pane)) pane.Dispose();
    }

    /// <summary>Close panes whose anchor is no longer a t: cell.</summary>
    public void Sync(IReadOnlySet<(int row, int col)> liveAnchors)
    {
        if (_panes.Count == 0) return;
        List<(int, int)>? dead = null;
        foreach (var key in _panes.Keys)
            if (!liveAnchors.Contains(key)) (dead ??= new()).Add(key);
        if (dead == null) return;
        foreach (var key in dead)
            if (_panes.Remove(key, out var pane)) pane.Dispose();
    }

    public bool AnyNeedsRedraw()
    {
        bool any = false;
        foreach (var p in _panes.Values) any |= p.NeedsRedraw;
        return any;
    }

    public void Dispose()
    {
        foreach (var p in _panes.Values) p.Dispose();
        _panes.Clear();
    }
}
