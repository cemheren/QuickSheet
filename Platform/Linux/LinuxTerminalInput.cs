using System.Runtime.InteropServices;
using System.Text;
using ExcelConsole.Terminal;
using static ExcelConsole.Platform.Linux.X11Methods;

namespace ExcelConsole.Platform.Linux;

/// <summary>Translates X11 key events into terminal key presses for a focused t: pane.</summary>
internal static class LinuxTerminalInput
{
    private const uint Mod4Mask = 1 << 6; // Super

    public const ulong XK_grave = 0x0060;
    public const ulong XK_Page_Up = 0xff55;
    public const ulong XK_Page_Down = 0xff56;
    public const ulong XK_Insert = 0xff63;
    public const ulong XK_ISO_Left_Tab = 0xfe20;
    public const ulong XK_V = 0x0056;

    /// <summary>Send one key press to the pane. Returns true if the key was handled.</summary>
    public static bool SendKey(TerminalPane pane, ref XKeyEvent ev)
    {
        ulong keysym = XLookupKeysym(ref ev, 0);           // unshifted
        TermMods mods = Mods(ev.state);
        string text = LookupText(ref ev);
        TermKey key = Map(keysym);

        if (keysym == XK_ISO_Left_Tab) key = TermKey.Tab;   // Shift+Tab on some layouts

        uint unshifted = keysym is >= 0x20 and <= 0x7e or >= 0xa0 and <= 0xff ? (uint)keysym
            : (keysym & 0xff000000) == 0x01000000 ? (uint)(keysym & 0x00ffffff) : 0;

        if (key == TermKey.Unidentified && text.Length == 0) return false;

        // Shift that produced the text (e.g. "A", "!") is consumed by the layout.
        TermMods consumed = text.Length > 0 ? mods & TermMods.Shift : TermMods.None;
        return pane.SendKey(key, mods, text, unshifted, consumed);
    }

    public static TermMods Mods(uint state)
    {
        TermMods m = TermMods.None;
        if ((state & ShiftMask) != 0) m |= TermMods.Shift;
        if ((state & ControlMask) != 0) m |= TermMods.Ctrl;
        if ((state & Mod1Mask) != 0) m |= TermMods.Alt;
        if ((state & Mod4Mask) != 0) m |= TermMods.Super;
        return m;
    }

    /// <summary>
    /// The text the key types with the current layout and Shift, ignoring Ctrl/Alt (the encoder
    /// derives Ctrl/Alt sequences itself). Empty for keys that type control characters.
    /// </summary>
    private static string LookupText(ref XKeyEvent ev)
    {
        var plain = ev;
        plain.state &= ~(ControlMask | Mod1Mask | Mod4Mask);
        IntPtr buf = Marshal.AllocHGlobal(32);
        try
        {
            int len = XLookupString(ref plain, buf, 32, out _, IntPtr.Zero);
            if (len <= 0) return "";
            var bytes = new byte[len];
            Marshal.Copy(buf, bytes, 0, len);
            // XLookupString returns Latin-1 for the default (non-UTF-8) locale; QuickSheet
            // never calls setlocale, so decode as Latin-1 unless it is valid multi-byte UTF-8.
            string s;
            try { s = new UTF8Encoding(false, true).GetString(bytes); }
            catch (DecoderFallbackException) { s = Encoding.Latin1.GetString(bytes); }
            foreach (char c in s)
                if (c < 0x20 || c == 0x7f) return "";
            return s;
        }
        finally { Marshal.FreeHGlobal(buf); }
    }

    public static TermKey Map(ulong keysym)
    {
        switch (keysym)
        {
            case XK_Return: return TermKey.Enter;
            case XK_Escape: return TermKey.Escape;
            case XK_Tab: return TermKey.Tab;
            case XK_BackSpace: return TermKey.Backspace;
            case XK_Delete: return TermKey.Delete;
            case XK_Insert: return TermKey.Insert;
            case XK_Home: return TermKey.Home;
            case XK_End: return TermKey.End;
            case XK_Page_Up: return TermKey.PageUp;
            case XK_Page_Down: return TermKey.PageDown;
            case XK_Left: return TermKey.ArrowLeft;
            case XK_Right: return TermKey.ArrowRight;
            case XK_Up: return TermKey.ArrowUp;
            case XK_Down: return TermKey.ArrowDown;
            case 0x20: return TermKey.Space;
            case 0x60: return TermKey.Backquote;
            case 0x5c: return TermKey.Backslash;
            case 0x5b: return TermKey.BracketLeft;
            case 0x5d: return TermKey.BracketRight;
            case 0x2c: return TermKey.Comma;
            case 0x3d: return TermKey.Equal;
            case 0x2d: return TermKey.Minus;
            case 0x2e: return TermKey.Period;
            case 0x27: return TermKey.Quote;
            case 0x3b: return TermKey.Semicolon;
            case 0x2f: return TermKey.Slash;
            case 0xff8d: return TermKey.NumpadEnter;      // KP_Enter
            case 0xffab: return TermKey.NumpadAdd;
            case 0xffad: return TermKey.NumpadSubtract;
            case 0xffaa: return TermKey.NumpadMultiply;
            case 0xffaf: return TermKey.NumpadDivide;
            case 0xffae: return TermKey.NumpadDecimal;
        }
        if (keysym is >= 0x61 and <= 0x7a) return TermKey.A + (int)(keysym - 0x61);
        if (keysym is >= 0x41 and <= 0x5a) return TermKey.A + (int)(keysym - 0x41);
        if (keysym is >= 0x30 and <= 0x39) return TermKey.Digit0 + (int)(keysym - 0x30);
        if (keysym is >= 0xffb0 and <= 0xffb9) return TermKey.Numpad0 + (int)(keysym - 0xffb0);
        if (keysym is >= 0xffbe and <= 0xffc9) return TermKey.F1 + (int)(keysym - 0xffbe);
        return TermKey.Unidentified;
    }
}
