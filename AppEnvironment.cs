namespace ExcelConsole;

/// <summary>
/// Exposes environment-level settings checked once at startup.
/// Respects the NO_COLOR standard (https://no-color.org/).
/// </summary>
public static class AppEnvironment
{
    /// <summary>
    /// True when the NO_COLOR environment variable is set (to any value),
    /// indicating that the program should suppress color/style output.
    /// </summary>
    public static bool NoColor { get; } =
        Environment.GetEnvironmentVariable("NO_COLOR") is not null;
}
