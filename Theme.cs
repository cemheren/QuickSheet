namespace ExcelConsole;

/// <summary>
/// Console color theme presets for the TUI spreadsheet.
/// Cycle themes with Ctrl+T.
/// </summary>
public class Theme
{
    public string Name { get; init; } = "Dark";
    public ConsoleColor Background { get; init; } = ConsoleColor.Black;
    public ConsoleColor Foreground { get; init; } = ConsoleColor.White;
    public ConsoleColor HeaderHighlight { get; init; } = ConsoleColor.DarkGray;
    public ConsoleColor SelectionBg { get; init; } = ConsoleColor.DarkGray;
    public ConsoleColor SelectionFg { get; init; } = ConsoleColor.White;
    public ConsoleColor SearchMatchBg { get; init; } = ConsoleColor.DarkYellow;
    public ConsoleColor SearchMatchFg { get; init; } = ConsoleColor.Black;
    public ConsoleColor SearchSelectedBg { get; init; } = ConsoleColor.Green;
    public ConsoleColor SearchSelectedFg { get; init; } = ConsoleColor.Black;
    public ConsoleColor StatusBarBg { get; init; } = ConsoleColor.White;
    public ConsoleColor StatusBarFg { get; init; } = ConsoleColor.Black;

    // Optional true-color overrides used only by the desktop renderers (WinForms / X11+Xft).
    // Console/TUI mode always falls back to the 16-color ConsoleColor fields above.
    // null = let the desktop renderer use ConsoleColorToRgb on the matching slot.
    public (int r, int g, int b)? BackgroundRgb { get; init; }
    public (int r, int g, int b)? ForegroundRgb { get; init; }
    public (int r, int g, int b)? HeaderHighlightRgb { get; init; }
    public (int r, int g, int b)? SelectionBgRgb { get; init; }
    public (int r, int g, int b)? SelectionFgRgb { get; init; }
    public (int r, int g, int b)? SearchMatchBgRgb { get; init; }
    public (int r, int g, int b)? SearchMatchFgRgb { get; init; }
    public (int r, int g, int b)? SearchSelectedBgRgb { get; init; }
    public (int r, int g, int b)? SearchSelectedFgRgb { get; init; }
    public (int r, int g, int b)? StatusBarBgRgb { get; init; }
    public (int r, int g, int b)? StatusBarFgRgb { get; init; }

    public static readonly Theme[] Presets =
    [
        new Theme
        {
            Name = "Dark",
            Background = ConsoleColor.Black,
            Foreground = ConsoleColor.White,
            HeaderHighlight = ConsoleColor.DarkGray,
            SelectionBg = ConsoleColor.DarkGray,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.DarkYellow,
            SearchMatchFg = ConsoleColor.Black,
            SearchSelectedBg = ConsoleColor.Green,
            SearchSelectedFg = ConsoleColor.Black,
            StatusBarBg = ConsoleColor.White,
            StatusBarFg = ConsoleColor.Black,
        },
        new Theme
        {
            Name = "Light",
            Background = ConsoleColor.White,
            Foreground = ConsoleColor.Black,
            HeaderHighlight = ConsoleColor.Gray,
            SelectionBg = ConsoleColor.DarkCyan,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.Yellow,
            SearchMatchFg = ConsoleColor.Black,
            SearchSelectedBg = ConsoleColor.Green,
            SearchSelectedFg = ConsoleColor.White,
            StatusBarBg = ConsoleColor.DarkBlue,
            StatusBarFg = ConsoleColor.White,
        },
        new Theme
        {
            Name = "Nord",
            Background = ConsoleColor.DarkBlue,
            Foreground = ConsoleColor.White,
            HeaderHighlight = ConsoleColor.DarkCyan,
            SelectionBg = ConsoleColor.DarkCyan,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.DarkYellow,
            SearchMatchFg = ConsoleColor.Black,
            SearchSelectedBg = ConsoleColor.Green,
            SearchSelectedFg = ConsoleColor.Black,
            StatusBarBg = ConsoleColor.Cyan,
            StatusBarFg = ConsoleColor.Black,
        },
        new Theme
        {
            Name = "Solarized",
            Background = ConsoleColor.DarkBlue,
            Foreground = ConsoleColor.Gray,
            HeaderHighlight = ConsoleColor.DarkGreen,
            SelectionBg = ConsoleColor.DarkGreen,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.DarkYellow,
            SearchMatchFg = ConsoleColor.Black,
            SearchSelectedBg = ConsoleColor.DarkMagenta,
            SearchSelectedFg = ConsoleColor.White,
            StatusBarBg = ConsoleColor.DarkCyan,
            StatusBarFg = ConsoleColor.White,
        },
        new Theme
        {
            // Solarized Light — Ethan Schoonover's cream-on-slate light palette.
            // ConsoleColor fields are the 16-color TUI fallback; *Rgb overrides land
            // the actual Solarized hex on the desktop renderers (WinForms / X11+Xft).
            Name = "SolarizedLight",
            Background = ConsoleColor.White,
            Foreground = ConsoleColor.DarkGray,
            HeaderHighlight = ConsoleColor.Gray,
            SelectionBg = ConsoleColor.DarkCyan,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.DarkYellow,
            SearchMatchFg = ConsoleColor.Black,
            SearchSelectedBg = ConsoleColor.DarkMagenta,
            SearchSelectedFg = ConsoleColor.White,
            StatusBarBg = ConsoleColor.DarkCyan,
            StatusBarFg = ConsoleColor.White,
            // True Solarized hex per https://ethanschoonover.com/solarized/
            BackgroundRgb       = (253, 246, 227), // base3   #fdf6e3
            ForegroundRgb       = (101, 123, 131), // base00  #657b83
            HeaderHighlightRgb  = (147, 161, 161), // base1   #93a1a1
            SelectionBgRgb      = (238, 232, 213), // base2   #eee8d5
            SelectionFgRgb      = ( 88, 110, 117), // base01  #586e75
            SearchMatchBgRgb    = (181, 137,   0), // yellow  #b58900
            SearchMatchFgRgb    = (253, 246, 227), // base3
            SearchSelectedBgRgb = (203,  75,  22), // orange  #cb4b16
            SearchSelectedFgRgb = (253, 246, 227), // base3
            StatusBarBgRgb      = (  7,  54,  66), // base02  #073642
            StatusBarFgRgb      = (147, 161, 161), // base1
        },
        new Theme
        {
            Name = "Matrix",
            Background = ConsoleColor.Black,
            Foreground = ConsoleColor.Green,
            HeaderHighlight = ConsoleColor.DarkGreen,
            SelectionBg = ConsoleColor.DarkGreen,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.DarkYellow,
            SearchMatchFg = ConsoleColor.Black,
            SearchSelectedBg = ConsoleColor.Cyan,
            SearchSelectedFg = ConsoleColor.Black,
            StatusBarBg = ConsoleColor.DarkGreen,
            StatusBarFg = ConsoleColor.White,
        },
        new Theme
        {
            // Inspired by the popular Dracula palette (draculatheme.com).
            Name = "Dracula",
            Background = ConsoleColor.Black,
            Foreground = ConsoleColor.White,
            HeaderHighlight = ConsoleColor.DarkMagenta,
            SelectionBg = ConsoleColor.DarkMagenta,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.Yellow,
            SearchMatchFg = ConsoleColor.Black,
            SearchSelectedBg = ConsoleColor.Magenta,
            SearchSelectedFg = ConsoleColor.Black,
            StatusBarBg = ConsoleColor.DarkMagenta,
            StatusBarFg = ConsoleColor.White,
        },
        new Theme
        {
            // Outrun-era magenta + cyan. Pair with a CRT and a sunset.
            Name = "Synthwave",
            Background = ConsoleColor.Black,
            Foreground = ConsoleColor.Magenta,
            HeaderHighlight = ConsoleColor.DarkMagenta,
            SelectionBg = ConsoleColor.DarkCyan,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.Yellow,
            SearchMatchFg = ConsoleColor.DarkMagenta,
            SearchSelectedBg = ConsoleColor.Cyan,
            SearchSelectedFg = ConsoleColor.Black,
            StatusBarBg = ConsoleColor.Magenta,
            StatusBarFg = ConsoleColor.Black,
        },
        new Theme
        {
            // Warm earth tones, Gruvbox-inspired.
            Name = "Gruvbox",
            Background = ConsoleColor.Black,
            Foreground = ConsoleColor.Yellow,
            HeaderHighlight = ConsoleColor.DarkYellow,
            SelectionBg = ConsoleColor.DarkYellow,
            SelectionFg = ConsoleColor.Black,
            SearchMatchBg = ConsoleColor.DarkRed,
            SearchMatchFg = ConsoleColor.White,
            SearchSelectedBg = ConsoleColor.Red,
            SearchSelectedFg = ConsoleColor.White,
            StatusBarBg = ConsoleColor.DarkYellow,
            StatusBarFg = ConsoleColor.Black,
        },
        new Theme
        {
            // Monokai-inspired: dark backdrop with green + magenta accents.
            Name = "Monokai",
            Background = ConsoleColor.Black,
            Foreground = ConsoleColor.White,
            HeaderHighlight = ConsoleColor.DarkGreen,
            SelectionBg = ConsoleColor.DarkMagenta,
            SelectionFg = ConsoleColor.White,
            SearchMatchBg = ConsoleColor.Yellow,
            SearchMatchFg = ConsoleColor.Black,
            SearchSelectedBg = ConsoleColor.Green,
            SearchSelectedFg = ConsoleColor.Black,
            StatusBarBg = ConsoleColor.DarkGreen,
            StatusBarFg = ConsoleColor.White,
        },
        new Theme
        {
            // Hotdog Stand — the legendary Windows 3.1 cursed palette.
            // Yes, on purpose. Cycle past it quickly. Or don't.
            Name = "HotdogStand",
            Background = ConsoleColor.Red,
            Foreground = ConsoleColor.Yellow,
            HeaderHighlight = ConsoleColor.Yellow,
            SelectionBg = ConsoleColor.Yellow,
            SelectionFg = ConsoleColor.Red,
            SearchMatchBg = ConsoleColor.Black,
            SearchMatchFg = ConsoleColor.Yellow,
            SearchSelectedBg = ConsoleColor.White,
            SearchSelectedFg = ConsoleColor.Red,
            StatusBarBg = ConsoleColor.Yellow,
            StatusBarFg = ConsoleColor.Black,
        },
    ];

    private static int _currentIndex = 0;
    public static Theme Current => Presets[_currentIndex];

    public static void CycleNext()
    {
        _currentIndex = (_currentIndex + 1) % Presets.Length;
    }

    /// <summary>
    /// Set the active theme by preset name (case-insensitive). No-op if name doesn't match.
    /// Used by `config: theme=...` cells to restore the theme on load.
    /// </summary>
    public static void SetByName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return;
        for (int i = 0; i < Presets.Length; i++)
        {
            if (Presets[i].Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                _currentIndex = i;
                return;
            }
        }
    }
}
