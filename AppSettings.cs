namespace ExcelConsole;

/// <summary>
/// Global application settings derived from environment variables and CLI flags.
/// </summary>
public static class AppSettings
{
    private static bool? _noColorOverride;

    /// <summary>
    /// True when the NO_COLOR environment variable is set (any value),
    /// per the https://no-color.org/ standard, or when --no-color is passed.
    /// When true, all terminal output must suppress ANSI color/style escape
    /// sequences. Desktop-mode platform rendering (WinForms, X11) is unaffected.
    /// </summary>
    public static bool NoColor =>
        _noColorOverride ?? Environment.GetEnvironmentVariable("NO_COLOR") != null;

    /// <summary>
    /// Called at startup when --no-color is passed on the command line.
    /// </summary>
    public static void ForceNoColor() => _noColorOverride = true;
}
