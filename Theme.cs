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
