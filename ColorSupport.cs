namespace ExcelConsole;

/// <summary>
/// Respects the NO_COLOR standard (https://no-color.org/).
/// When NO_COLOR is set (any value), color output should be suppressed.
/// </summary>
public static class ColorSupport
{
    public static bool Enabled { get; } =
        Environment.GetEnvironmentVariable("NO_COLOR") is null;
}
