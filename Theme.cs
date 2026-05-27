namespace ExcelConsole;

/// <summary>
/// Color theme presets for the desktop renderers. RGB hex is the source of truth.
/// Cycle themes with Ctrl+T.
/// </summary>
public class Theme
{
    public string Name { get; init; } = "Dark";

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
    /// <summary>Alternate row background for zebra striping (even rows).</summary>
    public (int r, int g, int b) ZebraStripeBgRgb { get; init; } = (18, 18, 18);

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
            ZebraStripeBgRgb    = ( 18,  18,  18),
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
            ZebraStripeBgRgb    = (225, 225, 225),
        },
        new Theme
        {
            // Nord — https://www.nordtheme.com
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
            ZebraStripeBgRgb    = ( 54,  60,  72),  // nord0.5
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
            SearchMatchBgRgb    = (203, 173,  77),  // softened yellow
            SearchMatchFgRgb    = (  0,  43,  54),
            SearchSelectedBgRgb = (213, 122,  84),  // softened orange
            SearchSelectedFgRgb = (253, 246, 227),
            StatusBarBgRgb      = ( 95, 183, 175),  // softened cyan
            StatusBarFgRgb      = (  0,  43,  54),
            ZebraStripeBgRgb    = (  7,  54,  66),  // base02
        },
        new Theme
        {
            // Solarized Light — matches VSCode's solarized-light-color-theme.json
            // (microsoft/vscode/extensions/theme-solarized-light). Status bar is the
            // light base2, not dark base02 — matching VSCode's actual look.
            Name = "SolarizedLight",
            BackgroundRgb       = (253, 246, 227),  // base3   #fdf6e3
            ForegroundRgb       = (101, 123, 131),  // base00  #657b83
            HeaderHighlightRgb  = (238, 232, 213),  // base2   #eee8d5 (editorWidget bg)
            SelectionBgRgb      = (238, 232, 213),  // base2   #eee8d5
            SelectionFgRgb      = (101, 123, 131),  // base00
            SearchMatchBgRgb    = (229, 229, 161),  // VSCode findMatchHighlight #e5e5a1
            SearchMatchFgRgb    = (101, 123, 131),  // base00
            SearchSelectedBgRgb = (245, 164,  15),  // VSCode findMatch #f5a40f
            SearchSelectedFgRgb = (253, 246, 227),  // base3
            StatusBarBgRgb      = (238, 232, 213),  // base2   #eee8d5 (VSCode statusBar bg)
            StatusBarFgRgb      = (101, 123, 131),  // base00  (VSCode statusBar fg)
            ZebraStripeBgRgb    = (238, 232, 213),  // base2
        },
        new Theme
        {
            Name = "Matrix",
            BackgroundRgb       = (  0,   0,   0),
            ForegroundRgb       = ( 90, 200, 110),  // softer phosphor green
            HeaderHighlightRgb  = ( 30,  80,  40),
            SelectionBgRgb      = ( 30,  80,  40),
            SelectionFgRgb      = (240, 240, 240),
            SearchMatchBgRgb    = (170, 170,  80),  // softened olive
            SearchMatchFgRgb    = (  0,   0,   0),
            SearchSelectedBgRgb = ( 90, 200, 200),  // softened cyan
            SearchSelectedFgRgb = (  0,   0,   0),
            StatusBarBgRgb      = ( 30,  80,  40),
            StatusBarFgRgb      = (240, 240, 240),
            ZebraStripeBgRgb    = ( 10,  14,  10),
        },
        new Theme
        {
            // Dracula — https://draculatheme.com
            Name = "Dracula",
            BackgroundRgb       = ( 40,  42,  54),  // bg           #282a36
            ForegroundRgb       = (248, 248, 242),  // fg           #f8f8f2
            HeaderHighlightRgb  = ( 68,  71,  90),  // current line #44475a
            SelectionBgRgb      = ( 68,  71,  90),
            SelectionFgRgb      = (248, 248, 242),
            SearchMatchBgRgb    = (241, 250, 140),  // yellow #f1fa8c
            SearchMatchFgRgb    = ( 40,  42,  54),
            SearchSelectedBgRgb = (255, 121, 198),  // pink   #ff79c6
            SearchSelectedFgRgb = ( 40,  42,  54),
            StatusBarBgRgb      = (189, 147, 249),  // purple #bd93f9
            StatusBarFgRgb      = ( 40,  42,  54),
            ZebraStripeBgRgb    = ( 48,  50,  62),  // slightly lighter bg
        },
        new Theme
        {
            // Synthwave — outrun-era magenta + cyan
            Name = "Synthwave",
            BackgroundRgb       = ( 30,  20,  50),
            ForegroundRgb       = (255, 140, 220),  // softened pink
            HeaderHighlightRgb  = (120,  50, 130),
            SelectionBgRgb      = ( 90, 170, 180),
            SelectionFgRgb      = (240, 240, 240),
            SearchMatchBgRgb    = (240, 220, 130),
            SearchMatchFgRgb    = (120,  50, 130),
            SearchSelectedBgRgb = (130, 220, 220),
            SearchSelectedFgRgb = (  0,   0,   0),
            StatusBarBgRgb      = (220, 130, 200),  // softened magenta
            StatusBarFgRgb      = (  0,   0,   0),
            ZebraStripeBgRgb    = ( 38,  26,  58),
        },
        new Theme
        {
            // Gruvbox — warm earth tones
            Name = "Gruvbox",
            BackgroundRgb       = ( 40,  40,  40),  // bg0   #282828
            ForegroundRgb       = (235, 219, 178),  // fg    #ebdbb2
            HeaderHighlightRgb  = ( 80,  73,  69),  // bg1   #504945
            SelectionBgRgb      = (220, 175,  90),  // softened yellow
            SelectionFgRgb      = ( 40,  40,  40),
            SearchMatchBgRgb    = (215, 110, 100),  // softened red
            SearchMatchFgRgb    = (235, 219, 178),
            SearchSelectedBgRgb = (250, 140, 120),  // softened bright red
            SearchSelectedFgRgb = ( 40,  40,  40),
            StatusBarBgRgb      = (220, 175,  90),
            StatusBarFgRgb      = ( 40,  40,  40),
            ZebraStripeBgRgb    = ( 50,  48,  47),  // bg0_s
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
            SearchMatchBgRgb    = (230, 219, 140),  // softened yellow
            SearchMatchFgRgb    = ( 39,  40,  34),
            SearchSelectedBgRgb = (180, 220,  90),  // softened green
            SearchSelectedFgRgb = ( 39,  40,  34),
            StatusBarBgRgb      = (235, 110, 150),  // softened pink
            StatusBarFgRgb      = (248, 248, 242),
            ZebraStripeBgRgb    = ( 47,  48,  42),
        },
        new Theme
        {
            // Hotdog Stand — Windows 3.1 cursed palette, softened to pastel
            // (salmon + butter) so it's a vibe rather than an eyesore.
            Name = "HotdogStand",
            BackgroundRgb       = (242, 180, 180),  // pastel salmon
            ForegroundRgb       = (115,  60,  60),  // muted brick text
            HeaderHighlightRgb  = (250, 230, 170),  // pastel butter
            SelectionBgRgb      = (250, 230, 170),
            SelectionFgRgb      = (115,  60,  60),
            SearchMatchBgRgb    = (235, 195, 200),
            SearchMatchFgRgb    = ( 90,  40,  40),
            SearchSelectedBgRgb = (255, 245, 220),
            SearchSelectedFgRgb = (180,  60,  60),
            StatusBarBgRgb      = (250, 230, 170),
            StatusBarFgRgb      = (115,  60,  60),
            ZebraStripeBgRgb    = (230, 168, 168),
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
