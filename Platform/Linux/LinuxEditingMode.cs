namespace ExcelConsole.Platform.Linux;

/// <summary>
/// Cell editing state for the Linux X11 host. Snapshots the target cell on Enter() so
/// commit always writes back to the correct cell even if selection moves.
/// Mirrors Platform/Windows/EditingMode.cs but without any System.Windows.Forms dependency.
/// </summary>
internal class LinuxEditingMode
{
    private bool _active;
    private string _editText = "";
    private int _cursorPos;
    private int _editRow;
    private int _editCol;

    public bool IsActive() => _active;
    public string EditText => _editText;
    public int CursorPosition => _cursorPos;
    public int EditRow => _editRow;
    public int EditCol => _editCol;

    public void Enter(GridManager grid, int row, int col)
    {
        _active = true;
        _editRow = row;
        _editCol = col;
        _editText = grid.GetCellValue(row, col);
        _cursorPos = _editText.Length;
    }

    public void Exit()
    {
        _active = false;
        _editText = "";
        _cursorPos = 0;
    }

    public void Commit(GridManager grid)
    {
        if (!_active) return;
        grid.SetCellValue(_editRow, _editCol, _editText);
        Exit();
    }

    public void MoveLeft() { if (_cursorPos > 0) _cursorPos--; }
    public void MoveRight() { if (_cursorPos < _editText.Length) _cursorPos++; }
    public void MoveHome() { _cursorPos = 0; }
    public void MoveEnd() { _cursorPos = _editText.Length; }

    public void Backspace()
    {
        if (_cursorPos > 0)
        {
            _editText = _editText.Remove(_cursorPos - 1, 1);
            _cursorPos--;
        }
    }

    public void DeleteForward()
    {
        if (_cursorPos < _editText.Length)
            _editText = _editText.Remove(_cursorPos, 1);
    }

    public void Insert(string s)
    {
        _editText = _editText.Insert(_cursorPos, s);
        _cursorPos += s.Length;
    }

    public string GetStatusText()
    {
        string cursored = _editText.Insert(_cursorPos, "│");
        return $" Edit: {cursored}  (Enter=OK  Esc=Cancel)";
    }

    /// <summary>
    /// Display string for the cell while editing: edit text with a visible cursor character
    /// inserted at <see cref="CursorPosition"/>, padded/truncated to <paramref name="width"/>.
    /// </summary>
    public string GetCellDisplay(int width)
    {
        string withCursor = _editText.Insert(_cursorPos, "│");
        if (withCursor.Length >= width)
        {
            // Keep cursor visible: show a window ending at (or just past) the cursor.
            int end = System.Math.Min(_cursorPos + 1, withCursor.Length);
            int start = System.Math.Max(0, end - width);
            return withCursor.Substring(start, System.Math.Min(width, withCursor.Length - start));
        }
        return withCursor.PadRight(width);
    }
}
