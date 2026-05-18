namespace ExcelConsole;

/// <summary>
/// Persists user preferences (theme, grid dimensions) across sessions.
/// Stored as key=value lines in the same state directory as autosave.csv.
/// </summary>
internal static class AppConfig
{
    private static readonly string ConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "ExcelConsole", "config.ini");

    private static Dictionary<string, string> _values = new(StringComparer.OrdinalIgnoreCase);

    public static void Load()
    {
        if (!File.Exists(ConfigPath)) return;
        try
        {
            foreach (var line in File.ReadAllLines(ConfigPath))
            {
                var eq = line.IndexOf('=');
                if (eq > 0)
                    _values[line[..eq].Trim()] = line[(eq + 1)..].Trim();
            }
        }
        catch { /* ignore corrupt config */ }
    }

    private static void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
            File.WriteAllLines(ConfigPath, _values.Select(kv => $"{kv.Key}={kv.Value}"));
        }
        catch { /* ignore save failures */ }
    }

    public static string? Get(string key) =>
        _values.TryGetValue(key, out var v) ? v : null;

    public static void Set(string key, string value)
    {
        _values[key] = value;
        Save();
    }
}
