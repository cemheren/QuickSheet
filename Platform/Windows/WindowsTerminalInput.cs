using System.Windows.Forms;
using ExcelConsole.Terminal;

namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Translates WinForms key events into terminal key presses for a focused t: pane.
///
/// Windows splits a typed character into KeyDown (which key) and KeyPress (which character,
/// after layout, Shift, AltGr and dead keys). Keys that type text wait for KeyPress so layouts
/// work; everything else (arrows, Enter, Ctrl/Alt chords, F-keys) is sent from KeyDown.
/// </summary>
internal sealed class WindowsTerminalInput
{
    private TermKey _pendingKey = TermKey.Unidentified;
    private TermMods _pendingMods;
    private char _highSurrogate;

    /// <summary>
    /// Handle KeyDown for a focused pane. Returns true if the key was sent (the caller should then
    /// suppress KeyPress); false if the character will arrive through <see cref="KeyPress"/>.
    /// </summary>
    public bool KeyDown(TerminalPane pane, KeyEventArgs e)
    {
        TermKey key = Map(e.KeyCode);
        TermMods mods = Mods(e);
        bool altGr = e.Control && e.Alt;   // AltGr arrives as Ctrl+Alt

        if (IsTextKey(e.KeyCode) && (!e.Control && !e.Alt || altGr))
        {
            _pendingKey = key;
            _pendingMods = mods;
            return false;
        }

        _pendingKey = TermKey.Unidentified;
        if (key == TermKey.Unidentified) return false;

        // The unmodified character for Ctrl/Alt chords on letters and digits (Ctrl+C → "c").
        string text = "";
        uint unshifted = 0;
        if (key is >= TermKey.A and <= TermKey.Z)
        {
            char c = (char)('a' + (key - TermKey.A));
            unshifted = c;
            text = e.Shift ? char.ToUpperInvariant(c).ToString() : c.ToString();
        }
        else if (key is >= TermKey.Digit0 and <= TermKey.Digit9)
        {
            char c = (char)('0' + (key - TermKey.Digit0));
            unshifted = c;
            text = e.Shift ? "" : c.ToString();
        }
        TermMods consumed = text.Length > 0 ? mods & TermMods.Shift : TermMods.None;
        pane.SendKey(key, mods, text, unshifted, consumed);
        return true;
    }

    /// <summary>Handle KeyPress (a typed character) for a focused pane.</summary>
    public void KeyPress(TerminalPane pane, char ch)
    {
        if (char.IsHighSurrogate(ch)) { _highSurrogate = ch; return; }
        string text;
        if (char.IsLowSurrogate(ch) && _highSurrogate != 0)
        {
            text = new string(new[] { _highSurrogate, ch });
            _highSurrogate = '\0';
        }
        else
        {
            _highSurrogate = '\0';
            if (ch < 0x20 || ch == 0x7f) return;   // control characters were sent from KeyDown
            text = ch.ToString();
        }

        TermKey key = _pendingKey;
        TermMods mods = _pendingMods;
        _pendingKey = TermKey.Unidentified;
        _pendingMods = TermMods.None;
        // AltGr's Ctrl+Alt produced the character; it isn't a modifier for the program.
        if ((mods & (TermMods.Ctrl | TermMods.Alt)) == (TermMods.Ctrl | TermMods.Alt))
            mods &= ~(TermMods.Ctrl | TermMods.Alt);

        uint unshifted = key is >= TermKey.A and <= TermKey.Z ? (uint)('a' + (key - TermKey.A))
            : key is >= TermKey.Digit0 and <= TermKey.Digit9 ? (uint)('0' + (key - TermKey.Digit0))
            : key == TermKey.Space ? ' ' : 0u;
        pane.SendKey(key, mods, text, unshifted, mods & TermMods.Shift);
    }

    public static TermMods Mods(KeyEventArgs e)
    {
        TermMods m = TermMods.None;
        if (e.Shift) m |= TermMods.Shift;
        if (e.Control) m |= TermMods.Ctrl;
        if (e.Alt) m |= TermMods.Alt;
        return m;
    }

    private static bool IsTextKey(Keys k) =>
        k is >= Keys.A and <= Keys.Z or >= Keys.D0 and <= Keys.D9 or Keys.Space
          or >= Keys.NumPad0 and <= Keys.NumPad9 or Keys.Multiply or Keys.Add or Keys.Subtract or Keys.Decimal or Keys.Divide
          or Keys.Oemtilde or Keys.OemMinus or Keys.Oemplus or Keys.OemOpenBrackets or Keys.OemCloseBrackets
          or Keys.OemPipe or Keys.OemSemicolon or Keys.OemQuotes or Keys.Oemcomma or Keys.OemPeriod
          or Keys.OemQuestion or Keys.OemBackslash;

    public static TermKey Map(Keys k)
    {
        switch (k)
        {
            case Keys.Enter: return TermKey.Enter;
            case Keys.Escape: return TermKey.Escape;
            case Keys.Tab: return TermKey.Tab;
            case Keys.Back: return TermKey.Backspace;
            case Keys.Delete: return TermKey.Delete;
            case Keys.Insert: return TermKey.Insert;
            case Keys.Home: return TermKey.Home;
            case Keys.End: return TermKey.End;
            case Keys.PageUp: return TermKey.PageUp;
            case Keys.PageDown: return TermKey.PageDown;
            case Keys.Left: return TermKey.ArrowLeft;
            case Keys.Right: return TermKey.ArrowRight;
            case Keys.Up: return TermKey.ArrowUp;
            case Keys.Down: return TermKey.ArrowDown;
            case Keys.Space: return TermKey.Space;
            case Keys.Oemtilde: return TermKey.Backquote;
            case Keys.OemMinus: return TermKey.Minus;
            case Keys.Oemplus: return TermKey.Equal;
            case Keys.OemOpenBrackets: return TermKey.BracketLeft;
            case Keys.OemCloseBrackets: return TermKey.BracketRight;
            case Keys.OemPipe: return TermKey.Backslash;
            case Keys.OemSemicolon: return TermKey.Semicolon;
            case Keys.OemQuotes: return TermKey.Quote;
            case Keys.Oemcomma: return TermKey.Comma;
            case Keys.OemPeriod: return TermKey.Period;
            case Keys.OemQuestion: return TermKey.Slash;
            case Keys.Multiply: return TermKey.NumpadMultiply;
            case Keys.Add: return TermKey.NumpadAdd;
            case Keys.Subtract: return TermKey.NumpadSubtract;
            case Keys.Decimal: return TermKey.NumpadDecimal;
            case Keys.Divide: return TermKey.NumpadDivide;
        }
        if (k is >= Keys.A and <= Keys.Z) return TermKey.A + (k - Keys.A);
        if (k is >= Keys.D0 and <= Keys.D9) return TermKey.Digit0 + (k - Keys.D0);
        if (k is >= Keys.NumPad0 and <= Keys.NumPad9) return TermKey.Numpad0 + (k - Keys.NumPad0);
        if (k is >= Keys.F1 and <= Keys.F12) return TermKey.F1 + (k - Keys.F1);
        return TermKey.Unidentified;
    }
}
