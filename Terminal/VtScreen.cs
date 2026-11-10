namespace ExcelConsole.Terminal;

internal static class VtScreen
{
    private static int _reported;

    /// <summary>Creates a libghostty-backed screen when the native library is loadable, otherwise LegacyVtScreen.</summary>
    public static IVtScreen Create(int cols, int rows, int maxScrollbackLines)
    {
        if (GhosttyNative.IsAvailable)
        {
            try
            {
                var s = new GhosttyVtScreen(cols, rows, maxScrollbackLines);
                Report($"[QuickSheet] terminal backend: libghostty-vt {GhosttyNative.Version}");
                return s;
            }
            catch (Exception ex)
            {
                Report($"[QuickSheet] terminal backend: legacy (libghostty-vt {GhosttyNative.Version} failed to create terminal: {ex.Message})");
            }
        }
        else
        {
            Report($"[QuickSheet] terminal backend: legacy (libghostty-vt unavailable: {GhosttyNative.LoadError})");
        }
        return new LegacyVtScreen(cols, rows, maxScrollbackLines);
    }

    /// <summary>True if libghostty-vt was loaded (evaluated once, lazily).</summary>
    public static bool UsingGhostty => GhosttyNative.IsAvailable;

    private static void Report(string msg)
    {
        if (Interlocked.Exchange(ref _reported, 1) != 0) return;
        try { Console.Error.WriteLine(msg); } catch { }
    }
}
