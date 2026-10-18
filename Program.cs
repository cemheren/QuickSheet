using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelConsole;

public class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        if (args.Contains("--help") || args.Contains("-h"))
        {
            PrintHelp();
            return;
        }

        if (args.Contains("--version") || args.Contains("-v"))
        {
            PrintVersion();
            return;
        }

        if (args.Contains("--list-extensions"))
        {
            PrintInstalledExtensions();
            return;
        }

        // Indices of arguments that are values for known flags (skip them when looking for CSV path).
        var flagValueIndices = new HashSet<int>();
        string[] valuedFlags = { "--export-md", "--export-html", "--export-json", "--grep" };
        foreach (var flag in valuedFlags)
        {
            int idx = Array.IndexOf(args, flag);
            if (idx >= 0 && idx + 1 < args.Length)
                flagValueIndices.Add(idx + 1);
        }
        string? csvPath = args
            .Where((a, i) => !a.StartsWith("--") && !flagValueIndices.Contains(i))
            .FirstOrDefault();

        int exportIdx = Array.IndexOf(args, "--export-md");
        if (exportIdx >= 0)
        {
            if (csvPath == null || exportIdx + 1 >= args.Length)
            {
                Console.Error.WriteLine("Usage: ExcelConsole <input.csv> --export-md <output.md>");
                Environment.Exit(2);
                return;
            }
            string outPath = args[exportIdx + 1];
            if (!File.Exists(csvPath))
            {
                Console.Error.WriteLine($"Input CSV not found: {csvPath}");
                Environment.Exit(1);
                return;
            }

            // Probe CSV for dimensions so the headless GridManager is big enough.
            var lines = ApplyGrepFilter(File.ReadAllLines(csvPath), args);
            int rows = Math.Max(1, lines.Length);
            int cols = 1;
            foreach (var line in lines)
            {
                int n = 1;
                bool inQuotes = false;
                foreach (char ch in line)
                {
                    if (ch == '"') inQuotes = !inQuotes;
                    else if (ch == ',' && !inQuotes) n++;
                }
                if (n > cols) cols = n;
            }
            // Constructor derives ColumnCount from (availableWidth - 4) / columnWidth (default 20).
            var grid = new GridManager(availableWidth: cols * 20 + 4, availableHeight: rows);
            grid.LoadFromCsvLines(lines);
            if (outPath == "-")
            {
                grid.WriteMarkdownTo(Console.Out);
            }
            else
            {
                grid.SaveToMarkdown(outPath);
                Console.WriteLine($"Wrote markdown: {outPath}");
            }
            return;
        }

        int htmlIdx = Array.IndexOf(args, "--export-html");
        if (htmlIdx >= 0)
        {
            if (csvPath == null || htmlIdx + 1 >= args.Length)
            {
                Console.Error.WriteLine("Usage: ExcelConsole <input.csv> --export-html <output.html>");
                Environment.Exit(2);
                return;
            }
            string htmlOut = args[htmlIdx + 1];
            if (!File.Exists(csvPath))
            {
                Console.Error.WriteLine($"Input CSV not found: {csvPath}");
                Environment.Exit(1);
                return;
            }

            var lines = ApplyGrepFilter(File.ReadAllLines(csvPath), args);
            int rows = Math.Max(1, lines.Length);
            int cols = 1;
            foreach (var line in lines)
            {
                int n = 1;
                bool inQuotes = false;
                foreach (char ch in line)
                {
                    if (ch == '"') inQuotes = !inQuotes;
                    else if (ch == ',' && !inQuotes) n++;
                }
                if (n > cols) cols = n;
            }
            var grid = new GridManager(availableWidth: cols * 20 + 4, availableHeight: rows);
            grid.LoadFromCsvLines(lines);
            if (htmlOut == "-")
            {
                grid.WriteHtmlTo(Console.Out);
            }
            else
            {
                grid.SaveToHtml(htmlOut);
                Console.WriteLine($"Wrote HTML: {htmlOut}");
            }
            return;
        }

        int jsonIdx = Array.IndexOf(args, "--export-json");
        if (jsonIdx >= 0)
        {
            if (csvPath == null || jsonIdx + 1 >= args.Length)
            {
                Console.Error.WriteLine("Usage: ExcelConsole <input.csv> --export-json <output.json>");
                Environment.Exit(2);
                return;
            }
            string jsonOut = args[jsonIdx + 1];
            if (!File.Exists(csvPath))
            {
                Console.Error.WriteLine($"Input CSV not found: {csvPath}");
                Environment.Exit(1);
                return;
            }

            var lines = ApplyGrepFilter(File.ReadAllLines(csvPath), args);
            int rows = Math.Max(1, lines.Length);
            int cols = 1;
            foreach (var line in lines)
            {
                int n = 1;
                bool inQuotes = false;
                foreach (char ch in line)
                {
                    if (ch == '"') inQuotes = !inQuotes;
                    else if (ch == ',' && !inQuotes) n++;
                }
                if (n > cols) cols = n;
            }
            var grid = new GridManager(availableWidth: cols * 20 + 4, availableHeight: rows);
            grid.LoadFromCsvLines(lines);
            if (jsonOut == "-")
            {
                grid.WriteJsonTo(Console.Out);
            }
            else
            {
                grid.SaveToJson(jsonOut);
                Console.WriteLine($"Wrote JSON: {jsonOut}");
            }
            return;
        }

#if PLATFORM_WINDOWS
        HideConsoleWindow();
        using var host = new ExcelConsole.Platform.Windows.WindowsDesktopHost();
        host.Run(csvPath);
#elif PLATFORM_LINUX
        string sessionType = Environment.GetEnvironmentVariable("XDG_SESSION_TYPE") ?? "";
        if (sessionType.Equals("wayland", StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine("Warning: QuickSheet requires an X11 session.");
            Console.Error.WriteLine("You are running Wayland. Select 'GNOME on Xorg' at the login screen,");
            Console.Error.WriteLine("or set DISPLAY and try with XWayland (may not work as true desktop layer).");
            Console.Error.WriteLine();
            Console.Error.WriteLine("Attempting via XWayland anyway...");
        }

        using var host = new ExcelConsole.Platform.Linux.LinuxDesktopHost();
        host.Run(csvPath);
#else
        Console.Error.WriteLine("QuickSheet is only supported on Windows and Linux (X11).");
#endif
    }

    private const string Version = "0.36.0";

    private static void PrintInstalledExtensions()
    {
        string root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".quicksheet", "extensions");
        if (!Directory.Exists(root))
        {
            Console.WriteLine("(no extensions installed)");
            Console.WriteLine($"Install one with `ext: github:user/repo` from inside QuickSheet.");
            Console.WriteLine($"Extensions live under {root}");
            return;
        }

        var dirs = Directory.GetDirectories(root).OrderBy(d => d).ToList();
        if (dirs.Count == 0)
        {
            Console.WriteLine("(no extensions installed)");
            return;
        }

        Console.WriteLine($"Installed extensions ({root}):");
        foreach (var dir in dirs)
        {
            string manifestPath = Path.Combine(dir, "quicksheet-extension.json");
            string name = Path.GetFileName(dir);
            if (File.Exists(manifestPath))
            {
                try
                {
                    using var doc = System.Text.Json.JsonDocument.Parse(File.ReadAllText(manifestPath));
                    string prefix = doc.RootElement.TryGetProperty("prefix", out var p) ? p.GetString() ?? "?" : "?";
                    string version = doc.RootElement.TryGetProperty("version", out var v) ? v.GetString() ?? "?" : "?";
                    Console.WriteLine($"  {prefix,-8} {name}  v{version}");
                }
                catch
                {
                    Console.WriteLine($"  ?        {name}  (bad manifest)");
                }
            }
            else
            {
                Console.WriteLine($"  ?        {name}  (no manifest)");
            }
        }
    }

    private static void PrintVersion()
    {
        string platform =
#if PLATFORM_WINDOWS
            "windows";
#elif PLATFORM_LINUX
            "linux";
#else
            "unknown";
#endif
        Console.WriteLine($"QuickSheet {Version} ({platform}, .NET {Environment.Version})");
    }

    private static void PrintHelp()
    {
        Console.WriteLine("QuickSheet — interactive spreadsheet as your desktop wallpaper");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  ExcelConsole [<file.csv>]                          Run as desktop wallpaper");
        Console.WriteLine("  ExcelConsole <file.csv> --export-md <out.md>       Headless: CSV → Markdown table (use - for stdout)");
        Console.WriteLine("  ExcelConsole <file.csv> --export-html <out.html>   Headless: CSV → styled HTML table (use - for stdout)");
        Console.WriteLine("  ExcelConsole <file.csv> --export-json <out.json>   Headless: CSV → JSON array of objects (use - for stdout)");
        Console.WriteLine("  ExcelConsole <file.csv> --grep <pattern> --export-md -  Filter rows by regex before export");
        Console.WriteLine("  ExcelConsole --help                                Show this help");
        Console.WriteLine("  ExcelConsole --version                             Show version");
        Console.WriteLine("  ExcelConsole --list-extensions                     List installed extensions");
        Console.WriteLine();
        Console.WriteLine("Cell prefixes:");
        Console.WriteLine("  r: <cmd>          Runnable command. Press Enter to launch.");
        Console.WriteLine("  i: <cmd>          Inline subprocess. Output streams back into the cell.");
        Console.WriteLine("  s: 1,2,3,...      Sparkline (unicode block bars). Also accepts range: s: A1::A10");
        Console.WriteLine("  c:<color>: <text> Cell color. Highlights background (red/green/blue/yellow/cyan/magenta/white/gray).");
        Console.WriteLine("  L: <cellRef>,<N>m Loop the target cell every N minutes.");
        Console.WriteLine("  ext: github:u/r   Install an extension repo (registers a new prefix).");
        Console.WriteLine("  http(s)://...     Hyperlink. Highlighted, opens in browser on Enter.");
        Console.WriteLine();
        Console.WriteLine("Range references work inside text: {A1::C10}");
        Console.WriteLine("Tour: docs/tour.md · Issues: github.com/cemheren/QuickSheet/issues");
    }

    private static string[] ApplyGrepFilter(string[] lines, string[] args)
    {
        int grepIdx = Array.IndexOf(args, "--grep");
        if (grepIdx < 0 || grepIdx + 1 >= args.Length)
            return lines;

        string pattern = args[grepIdx + 1];
        Regex regex;
        try
        {
            regex = new Regex(pattern, RegexOptions.IgnoreCase);
        }
        catch (RegexParseException)
        {
            Console.Error.WriteLine($"Invalid regex pattern: {pattern}");
            Environment.Exit(2);
            return lines;
        }

        var result = new List<string>();
        // Always keep the header row (first line).
        if (lines.Length > 0)
            result.Add(lines[0]);

        for (int i = 1; i < lines.Length; i++)
        {
            if (regex.IsMatch(lines[i]))
                result.Add(lines[i]);
        }

        return result.ToArray();
    }

#if PLATFORM_WINDOWS
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private static void HideConsoleWindow()
    {
        IntPtr console = GetConsoleWindow();
        if (console != IntPtr.Zero)
            ShowWindow(console, 0); // SW_HIDE
    }

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
#endif
}
