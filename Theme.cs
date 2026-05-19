namespace ExcelConsole;

/// <summary>
/// Color theme presets. RGB hex is the source of truth; the TUI mode derives
/// the nearest of the 16 ConsoleColors at render time. Cycle themes with Ctrl+T.
/// </summary>
public class Theme
{
    public string Name { get; init; } = "Dark";

    // Slot RGB. Desktop renderers (WinForms / X11+Xft) read these directly.
    public (int r, int g, int b) BackgroundRgb { get; init; } = (0, 0, 0);
    public (int r, int g, int b) ForegroundRgb { get; init; } = (240, 240, 240);
    public (int r, int g, int b) HeaderHighlightRgb { get; init; } = (64, 64, 64);
    public (int r, int g, int b) SelectionBgRgb { get; init; } = (64, 64, 64);
    public (int r, int g, int b) SelectionFgRgb { get; init; } = (240, 240, 240);
    public (int r, int g, int b) SearchMatchBgRgb { get; init; } = (139, 139, 0);
    public (int r, int g, int b) SearchMatchFgRgb { get; init; } = (0, 0, 0);
    public (int r, int g, int b) SearchSelectedBgRgb { get; init; } = (0, 200, 0);
    public (int r, int g, int b) SearchSelectedFgRgb { get; init; } = (0, 0, 0);
    public (int r, int g, int b) StatusBarBgRgb { get; init; } = (240, 240, 240);
    public (int r, int g, int b) StatusBarFgRgb { get; init; } = (0, 0, 0);

    // Derived ConsoleColor slots for the TUI (Console.BackgroundColor / ForegroundColor).
    // Maps to the nearest of the 16 named ConsoleColors by squared RGB distance.
    public ConsoleColor Background       => Nearest(BackgroundRgb);
    public ConsoleColor Foreground       => Nearest(ForegroundRgb);
    public ConsoleColor HeaderHighlight  => Nearest(HeaderHighlightRgb);
    public ConsoleColor SelectionBg      => Nearest(SelectionBgRgb);
    public ConsoleColor SelectionFg      => Nearest(SelectionFgRgb);
    public ConsoleColor SearchMatchBg    => Nearest(SearchMatchBgRgb);
    public ConsoleColor SearchMatchFg    => Nearest(SearchMatchFgRgb);
    public ConsoleColor SearchSelectedBg => Nearest(SearchSelectedBgRgb);
    public ConsoleColor SearchSelectedFg => Nearest(SearchSelectedFgRgb);
    public ConsoleColor StatusBarBg      => Nearest(StatusBarBgRgb);
    public ConsoleColor StatusBarFg      => Nearest(StatusBarFgRgb);

    public static readonly Theme[] Presets =
    [
        new Theme
        {
            Name = "Dark",
            BackgroundRgb       = (  0,   0,   0),
            ForegroundRgb       = (240, 240, 240),
            HeaderHighlightRgb  = ( 64,  64,  64),
            SelectionBgRgb      = ( 64,  64,  64),
            SelectionFgRgb      = (240, 240, 240),
            SearchMatchBgRgb    = (139, 139,   0),
            SearchMatchFgRgb    = (  0,   0,   0),
            SearchSelectedBgRgb = (  0, 200,   0),
            SearchSelectedFgRgb = (  0,   0,   0),
            StatusBarBgRgb      = (240, 240, 240),
            StatusBarFgRgb      = (  0,   0,   0),
        },
        new Theme
        {
            Name = "Light",
            BackgroundRgb       = (240, 240, 240),
            ForegroundRgb       = (  0,   0,   0),
            HeaderHighlightRgb  = (169, 169, 169),
            SelectionBgRgb      = (  0, 139, 139),
            SelectionFgRgb      = (240, 240, 240),
            SearchMatchBgRgb    = (200, 200,   0),
            SearchMatchFgRgb    = (  0,   0,   0),
            SearchSelectedBgRgb = (  0, 200,   0),
            SearchSelectedFgRgb = (240, 240, 240),
            StatusBarBgRgb      = (  0,   0, 139),
            StatusBarFgRgb      = (240, 240, 240),
        },
        new Theme
        {
            Name = "Nord",
            BackgroundRgb       = ( 46,  52,  64),  // nord0  #2e3440
            ForegroundRgb       = (216, 222, 233),  // nord4  #d8dee9
            HeaderHighlightRgb  = ( 67,  76,  94),  // nord1  #434c5e
            SelectionBgRgb      = ( 76,  86, 106),  // nord3  #4c566a
            SelectionFgRgb      = (236, 239, 244),  // nord6  #eceff4
            SearchMatchBgRgb    = (235, 203, 139),  // nord13 #ebcb8b
            SearchMatchFgRgb    = ( 46,  52,  64),
            SearchSelectedBgRgb = (163, 190, 140),  // nord14 #a3be8c
            SearchSelectedFgRgb = ( 46,  52,  64),
            StatusBarBgRgb      = (136, 192, 208),  // nord8  #88c0d0
            StatusBarFgRgb      = ( 46,  52,  64),
        },
        new Theme
        {
            // Solarized Dark — Ethan Schoonover, https://ethanschoonover.com/solarized/
            Name = "Solarized",
            BackgroundRgb       = (  0,  43,  54),  // base03 #002b36
            ForegroundRgb       = (131, 148, 150),  // base0  #839496
            HeaderHighlightRgb  = (  7,  54,  66),  // base02 #073642
            SelectionBgRgb      = (  7,  54,  66),
            SelectionFgRgb      = (147, 161, 161),  // base1  #93a1a1
            SearchMatchBgRgb    = (181, 137,   0),  // yellow #b58900
            SearchMatchFgRgb    = (  0,  43,  54),
            SearchSelectedBgRgb = (203,  75,  22),  // orange #cb4b16
            SearchSelectedFgRgb = (253, 246, 227),
            StatusBarBgRgb      = ( 42, 161, 152),  // cyan   #2aa198
            StatusBarFgRgb      = (  0,  43,  54),
        },
        new Theme
        {
            // Solarized Light — Schoonover. Cream background, slate text, VSCode vibe.
            Name = "SolarizedLight",
            BackgroundRgb       = (253, 246, 227),  // base3   #fdf6e3
            ForegroundRgb       = (101, 123, 131),  // base00  #657b83
            HeaderHighlightRgb  = (147, 161, 161),  // base1   #93a1a1
            SelectionBgRgb      = (238, 232, 213),  // base2   #eee8d5
            SelectionFgRgb      = ( 88, 110, 117),  // base01  #586e75
            SearchMatchBgRgb    = (181, 137,   0),  // yellow  #b58900
            SearchMatchFgRgb    = (253, 246, 227),
            SearchSelectedBgRgb = (203,  75,  22),  // orange  #cb4b16
            SearchSelectedFgRgb = (253, 246, 227),
            StatusBarBgRgb      = (  7,  54,  66),  // base02  #073642
            StatusBarFgRgb      = (147, 161, 161),  // base1
        },
        new Theme
        {
            Name = "Matrix",
            BackgroundRgb       = (  0,   0,   0),
            ForegroundRgb       = (  0, 200,   0),
            HeaderHighlightRgb  = (  0, 100,   0),
            SelectionBgRgb      = (  0, 100,   0),
            SelectionFgRgb      = (240, 240, 240),
            SearchMatchBgRgb    = (139, 139,   0),
            SearchMatchFgRgb    = (  0,   0,   0),
            SearchSelectedBgRgb = (  0, 200, 200),
            SearchSelectedFgRgb = (  0,   0,   0),
            StatusBarBgRgb      = (  0, 100,   0),
            StatusBarFgRgb      = (240, 240, 240),
        },
        new Theme
        {
            // Dracula — draculatheme.com
            Name = "Dracula",
            BackgroundRgb       = ( 40,  42,  54),  // bg            #282a36
            ForegroundRgb       = (248, 248, 242),  // fg            #f8f8f2
            HeaderHighlightRgb  = ( 68,  71,  90),  // current line  #44475a
            SelectionBgRgb      = ( 68,  71,  90),
            SelectionFgRgb      = (248, 248, 242),
            SearchMatchBgRgb    = (241, 250, 140),  // yellow #f1fa8c
            SearchMatchFgRgb    = ( 40,  42,  54),
            SearchSelectedBgRgb = (255, 121, 198),  // pink   #ff79c6
            SearchSelectedFgRgb = ( 40,  42,  54),
            StatusBarBgRgb      = (189, 147, 249),  // purple #bd93f9
            StatusBarFgRgb      = ( 40,  42,  54),
        },
        new Theme
        {
            // Synthwave — outrun-era magenta + cyan
            Name = "Synthwave",
            BackgroundRgb       = ( 26,  16,  41),
            ForegroundRgb       = (255,  64, 200),
            HeaderHighlightRgb  = (100,  20, 100),
            SelectionBgRgb      = (  0, 139, 139),
            SelectionFgRgb      = (240, 240, 240),
            SearchMatchBgRgb    = (255, 230,  80),
            SearchMatchFgRgb    = (100,  20, 100),
            SearchSelectedBgRgb = (  0, 230, 230),
            SearchSelectedFgRgb = (  0,   0,   0),
            StatusBarBgRgb      = (255,  64, 200),
            StatusBarFgRgb      = (  0,   0,   0),
        },
        new Theme
        {
            // Gruvbox — warm earth tones
            Name = "Gruvbox",
            BackgroundRgb       = ( 40,  40,  40),  // bg0    #282828
            ForegroundRgb       = (235, 219, 178),  // fg     #ebdbb2
            HeaderHighlightRgb  = ( 80,  73,  69),  // bg1    #504945
            SelectionBgRgb      = (215, 153,  33),  // yellow #d79921
            SelectionFgRgb      = ( 40,  40,  40),
            SearchMatchBgRgb    = (204,  36,  29),  // red    #cc241d
            SearchMatchFgRgb    = (235, 219, 178),
            SearchSelectedBgRgb = (251,  73,  52),  // bright red #fb4934
            SearchSelectedFgRgb = (235, 219, 178),
            StatusBarBgRgb      = (215, 153,  33),
            StatusBarFgRgb      = ( 40,  40,  40),
        },
        new Theme
        {
            // Monokai — Wimer Hazenberg
            Name = "Monokai",
            BackgroundRgb       = ( 39,  40,  34),  // bg      #272822
            ForegroundRgb       = (248, 248, 242),  // fg      #f8f8f2
            HeaderHighlightRgb  = (117, 113,  94),  // comment #75715e
            SelectionBgRgb      = ( 73,  72,  62),
            SelectionFgRgb      = (248, 248, 242),
            SearchMatchBgRgb    = (230, 219, 116),  // yellow  #e6db74
            SearchMatchFgRgb    = ( 39,  40,  34),
            SearchSelectedBgRgb = (166, 226,  46),  // green   #a6e22e
            SearchSelectedFgRgb = ( 39,  40,  34),
            StatusBarBgRgb      = (249,  38, 114),  // pink    #f92672
            StatusBarFgRgb      = (248, 248, 242),
        },
        new Theme
        {
            // Hotdog Stand — Windows 3.1 cursed palette. On purpose.
            Name = "HotdogStand",
            BackgroundRgb       = (200,   0,   0),
            ForegroundRgb       = (255, 255,   0),
            HeaderHighlightRgb  = (255, 255,   0),
            SelectionBgRgb      = (255, 255,   0),
            SelectionFgRgb      = (200,   0,   0),
            SearchMatchBgRgb    = (  0,   0,   0),
            SearchMatchFgRgb    = (255, 255,   0),
            SearchSelectedBgRgb = (255, 255, 255),
            SearchSelectedFgRgb = (200,   0,   0),
            StatusBarBgRgb      = (255, 255,   0),
            StatusBarFgRgb      = (  0,   0,   0),
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

    // ── ConsoleColor mapping for the TUI ──────────────────────────────
    //
    // System.Console only accepts the 16-entry ConsoleColor enum, so the TUI
    // path picks the closest one to each slot's true RGB. The reference RGB
    // values here match the desktop renderers' ConsoleColorToRgb mapping so
    // both surfaces stay visually consistent on the legacy 16-color palette.

    private static readonly (ConsoleColor cc, int r, int g, int b)[] _palette =
    [
        (ConsoleColor.Black,         0,   0,   0),
        (ConsoleColor.DarkBlue,      0,   0, 139),
        (ConsoleColor.DarkGreen,     0, 100,   0),
        (ConsoleColor.DarkCyan,      0, 139, 139),
        (ConsoleColor.DarkRed,     139,   0,   0),
        (ConsoleColor.DarkMagenta, 139,   0, 139),
        (ConsoleColor.DarkYellow,  139, 139,   0),
        (ConsoleColor.Gray,        169, 169, 169),
        (ConsoleColor.DarkGray,     64,  64,  64),
        (ConsoleColor.Blue,         30,  80, 200),
        (ConsoleColor.Green,         0, 200,   0),
        (ConsoleColor.Cyan,          0, 200, 200),
        (ConsoleColor.Red,         200,   0,   0),
        (ConsoleColor.Magenta,     200,   0, 200),
        (ConsoleColor.Yellow,      200, 200,   0),
        (ConsoleColor.White,       240, 240, 240),
    ];

    private static ConsoleColor Nearest((int r, int g, int b) rgb)
    {
        int bestDist = int.MaxValue;
        ConsoleColor best = ConsoleColor.White;
        foreach (var (cc, r, g, b) in _palette)
        {
            int dr = r - rgb.r, dg = g - rgb.g, db = b - rgb.b;
            int d = dr * dr + dg * dg + db * db;
            if (d < bestDist) { bestDist = d; best = cc; }
        }
        return best;
    }
}
