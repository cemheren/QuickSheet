using System.Runtime.InteropServices;
using System.Text;
using ExcelConsole.Features;
using static ExcelConsole.Platform.Linux.X11Methods;

namespace ExcelConsole.Platform.Linux;

/// <summary>
/// Cell editing state for the Linux X11 host. Snapshots target cell on Enter() so
/// commit always writes back to the correct cell even if selection moves.
/// Mirrors Platform/Windows/EditingMode.cs but with X11 keysym dispatch instead of WinForms.
/// </summary>
internal class LinuxEditingMode : IMode
{
    private bool _active;
    private string _editText = "";
    private int _cursorPos;
    private int _editRow;
    private int _editCol;

    public LinuxEditingMode(GridManager grid) { Grid = grid; }

    public GridManager Grid { get; set; }

    public bool IsActive() => _active;
    public string EditText => _editText;
    public int CursorPosition => _cursorPos;
    public int EditRow => _editRow;
    public int EditCol => _editCol;

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
        _editText = "";
        _cursorPos = 0;
    }

    public void Commit()
    {
        if (!_active) return;
        Grid.SetCellValue(_editRow, _editCol, _editText);
        Exit();
    }

    public void Insert(string s)
    {
        _editText = _editText.Insert(_cursorPos, s);
        _cursorPos += s.Length;
    }

    public bool HandleKeyEventLinux(ulong keysym, ref XKeyEvent keyEvent, bool ctrl)
    {
        if (!_active) return false;

        switch (keysym)
        {
            case XK_Return:
                Commit();
                return true;
            case XK_Escape:
                Exit();
                return true;
            case XK_Left:
                if (_cursorPos > 0) _cursorPos--;
                return true;
            case XK_Right:
                if (_cursorPos < _editText.Length) _cursorPos++;
                return true;
            case XK_Home:
                _cursorPos = 0;
                return true;
            case XK_End:
                _cursorPos = _editText.Length;
                return true;
            case XK_BackSpace:
                if (_cursorPos > 0)
                {
                    _editText = _editText.Remove(_cursorPos - 1, 1);
                    _cursorPos--;
                }
                return true;
            case XK_Delete:
                if (_cursorPos < _editText.Length)
                    _editText = _editText.Remove(_cursorPos, 1);
                return true;
            default:
                if (ctrl) return false;
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
                            Insert(ch);
                    }
                }
                finally { Marshal.FreeHGlobal(buf); }
                return true;
        }
    }

    public string GetStatusText()
    {
        string cursored = _editText.Insert(_cursorPos, "│");
        return $" Edit: {cursored}  (Enter=OK  Esc=Cancel)";
    }

    /// <summary>
    /// Display string for the cell while editing: edit text with visible cursor at
    /// <see cref="CursorPosition"/>, padded/truncated to <paramref name="width"/>.
    /// </summary>
    public string GetCellDisplay(int width)
    {
        string withCursor = _editText.Insert(_cursorPos, "│");
        if (withCursor.Length >= width)
        {
            int end = System.Math.Min(_cursorPos + 1, withCursor.Length);
            int start = System.Math.Max(0, end - width);
            return withCursor.Substring(start, System.Math.Min(width, withCursor.Length - start));
        }
        return withCursor.PadRight(width);
    }
}
