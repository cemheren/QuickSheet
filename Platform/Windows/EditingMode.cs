using System.Windows.Forms;
using ExcelConsole.Features;

namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Encapsulates cell-editing state and key handling for DesktopForm.
/// Snapshots the target cell on Enter() so commit always writes back
/// to the correct cell even if selection moves.
/// </summary>
internal class EditingMode : IMode
{
    private bool _active;
    private string _editText = "";
    private int _cursorPos;
    private int _editRow;
    private int _editCol;

    public EditingMode(GridManager grid)
    {
        Grid = grid;
    }

    public GridManager Grid { get; set; }

    public string EditText => _editText;
    public int CursorPosition => _cursorPos;

    /// <summary>Start editing the currently selected cell.</summary>
    public void Enter()
    {
        _active = true;
        (_editRow, _editCol) = Grid.GetCurrentCell();
        _editText = Grid.GetCellValue(_editRow, _editCol);
        _cursorPos = _editText.Length;
    }

    public void Exit()
    {
        _active = false;
    }

    public void Commit()
    {
        Grid.SetCellValue(_editRow, _editCol, _editText);
        _active = false;
    }

    public bool IsActive() => _active;

    public bool HandleKeyEvent(KeyEventArgs e)
    {
        if (!_active) return false;

        switch (e.KeyCode)
        {
            case Keys.Left:
                if (_cursorPos > 0) _cursorPos--;
                break;
            case Keys.Right:
                if (_cursorPos < _editText.Length) _cursorPos++;
                break;
            case Keys.Home:
                _cursorPos = 0;
                break;
            case Keys.End:
                _cursorPos = _editText.Length;
                break;
            case Keys.Back:
                if (_cursorPos > 0)
                {
                    _editText = _editText.Remove(_cursorPos - 1, 1);
                    _cursorPos--;
                }
                break;
            case Keys.Delete:
                if (_cursorPos < _editText.Length)
                    _editText = _editText.Remove(_cursorPos, 1);
                break;
            case Keys.V:
                if (e.Control && Clipboard.ContainsText())
                {
                    string clipText = Clipboard.GetText().Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                    _editText = _editText.Insert(_cursorPos, clipText);
                    _cursorPos += clipText.Length;
                }
                else
                    return false;
                break;
            case Keys.Enter:
                Commit();
                break;
            case Keys.Escape:
                Exit();
                break;
            default:
                return false;
        }
        return true;
    }

    /// <summary>Insert a printable character at cursor position.</summary>
    public void HandleCharInput(char c)
    {
        _editText = _editText.Insert(_cursorPos, c.ToString());
        _cursorPos++;
    }

    /// <summary>Status bar text while editing.</summary>
    public string GetStatusText()
    {
        var cursor = _editText.Insert(_cursorPos, "\u2502");
        return $" Edit: {cursor}  (Enter=OK  Esc=Cancel)";
    }
}
