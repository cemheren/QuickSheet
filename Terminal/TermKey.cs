namespace ExcelConsole.Terminal;

/// <summary>
/// Physical/logical keys for the terminal key encoder. Values are libghostty-vt's GhosttyKey
/// (include/ghostty/vt/key/event.h, vt-1.0.0), so they pass straight through to the encoder.
/// Platform code maps native key codes (X11 keysyms, WinForms Keys) to these.
/// </summary>
internal enum TermKey
{
    Unidentified = 0,
    Backquote = 1, Backslash = 2, BracketLeft = 3, BracketRight = 4, Comma = 5,
    Digit0 = 6, Digit1 = 7, Digit2 = 8, Digit3 = 9, Digit4 = 10,
    Digit5 = 11, Digit6 = 12, Digit7 = 13, Digit8 = 14, Digit9 = 15,
    Equal = 16,
    A = 20, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, // 20..45
    Minus = 46, Period = 47, Quote = 48, Semicolon = 49, Slash = 50,
    Backspace = 53,
    Enter = 58,
    Space = 63,
    Tab = 64,
    Delete = 68, End = 69, Home = 71, Insert = 72, PageDown = 73, PageUp = 74,
    ArrowDown = 75, ArrowLeft = 76, ArrowRight = 77, ArrowUp = 78,
    Numpad0 = 80, Numpad1, Numpad2, Numpad3, Numpad4, Numpad5, Numpad6, Numpad7, Numpad8, Numpad9, // 80..89
    NumpadAdd = 90, NumpadDecimal = 95, NumpadDivide = 96, NumpadEnter = 97,
    NumpadMultiply = 104, NumpadSubtract = 107,
    Escape = 120,
    F1 = 121, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12, // 121..132
}

[Flags]
internal enum TermMods : ushort
{
    None = 0,
    Shift = 1 << 0,
    Ctrl = 1 << 1,
    Alt = 1 << 2,
    Super = 1 << 3,
}
