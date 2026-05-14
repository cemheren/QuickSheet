namespace ExcelConsole;

/// <summary>
/// Tracks cell edits and supports undo (Ctrl+Z) / redo (Ctrl+Y).
/// Each "action" is a group of one or more cell changes that happened atomically
/// (e.g. a row delete touches many cells but is one undoable action).
/// </summary>
public class UndoManager
{
    private readonly record struct CellChange(int Row, int Col, string OldValue, string NewValue);

    private readonly List<List<CellChange>> _undoStack = new();
    private readonly List<List<CellChange>> _redoStack = new();
    private List<CellChange>? _pendingGroup;

    private const int MaxHistory = 200;

    /// <summary>Begin a group of related changes (e.g. a row shift).</summary>
    public void BeginGroup() => _pendingGroup = new List<CellChange>();

    /// <summary>Record a single cell change. Call between BeginGroup/EndGroup,
    /// or standalone (auto-wraps in its own group).</summary>
    public void RecordChange(int row, int col, string oldValue, string newValue)
    {
        if (oldValue == newValue) return;

        var change = new CellChange(row, col, oldValue, newValue);

        if (_pendingGroup != null)
        {
            _pendingGroup.Add(change);
        }
        else
        {
            _undoStack.Add(new List<CellChange> { change });
            TrimStack(_undoStack);
            _redoStack.Clear();
        }
    }

    /// <summary>End the current group and push it as one undoable action.</summary>
    public void EndGroup()
    {
        if (_pendingGroup is { Count: > 0 })
        {
            _undoStack.Add(_pendingGroup);
            TrimStack(_undoStack);
            _redoStack.Clear();
        }
        _pendingGroup = null;
    }

    /// <summary>Undo the last action group. Returns the cells to restore, or null if nothing to undo.</summary>
    public IReadOnlyList<(int Row, int Col, string Value)>? Undo()
    {
        if (_undoStack.Count == 0) return null;

        var group = _undoStack[^1];
        _undoStack.RemoveAt(_undoStack.Count - 1);
        _redoStack.Add(group);

        // Return old values in reverse order (undo row shifts correctly)
        var result = new List<(int, int, string)>(group.Count);
        for (int i = group.Count - 1; i >= 0; i--)
            result.Add((group[i].Row, group[i].Col, group[i].OldValue));
        return result;
    }

    /// <summary>Redo the last undone action group. Returns the cells to restore, or null if nothing to redo.</summary>
    public IReadOnlyList<(int Row, int Col, string Value)>? Redo()
    {
        if (_redoStack.Count == 0) return null;

        var group = _redoStack[^1];
        _redoStack.RemoveAt(_redoStack.Count - 1);
        _undoStack.Add(group);

        var result = new List<(int, int, string)>(group.Count);
        foreach (var c in group)
            result.Add((c.Row, c.Col, c.NewValue));
        return result;
    }

    public bool CanUndo => _undoStack.Count > 0;
    public bool CanRedo => _redoStack.Count > 0;

    private static void TrimStack(List<List<CellChange>> stack)
    {
        while (stack.Count > MaxHistory)
            stack.RemoveAt(0);
    }
}
