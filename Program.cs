using System.Runtime.InteropServices;
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

        string? csvPath = args.FirstOrDefault(a => !a.StartsWith("--"));

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
            var lines = File.ReadAllLines(csvPath);
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
            grid.LoadFromCsv(csvPath);
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

            var lines = File.ReadAllLines(csvPath);
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
            grid.LoadFromCsv(csvPath);
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

            var lines = File.ReadAllLines(csvPath);
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
            grid.LoadFromCsv(csvPath);
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

        int sortIdx = Array.IndexOf(args, "--sort");
        if (sortIdx >= 0)
        {
            if (csvPath == null || sortIdx + 1 >= args.Length)
            {
                Console.Error.WriteLine("Usage: ExcelConsole <input.csv> --sort <column> [--desc]");
                Console.Error.WriteLine("  <column> is a 0-based index or a header name.");
                Environment.Exit(2);
                return;
            }
            string sortCol = args[sortIdx + 1];
            bool descending = args.Contains("--desc");
            if (!File.Exists(csvPath))
            {
                Console.Error.WriteLine($"Input CSV not found: {csvPath}");
                Environment.Exit(1);
                return;
            }

            var allLines = File.ReadAllLines(csvPath);
            if (allLines.Length == 0)
            {
                return;
            }

            var parsed = allLines.Select(line => ParseCsvLine(line)).ToList();
            int colIdx;
            if (int.TryParse(sortCol, out int parsedIdx))
            {
                colIdx = parsedIdx;
            }
            else
            {
                // Match by header name (first row)
                colIdx = -1;
                for (int i = 0; i < parsed[0].Length; i++)
                {
                    if (string.Equals(parsed[0][i], sortCol, StringComparison.OrdinalIgnoreCase))
                    {
                        colIdx = i;
                        break;
                    }
                }
                if (colIdx < 0)
                {
                    Console.Error.WriteLine($"Column not found: {sortCol}");
                    Console.Error.WriteLine($"Available columns: {string.Join(", ", parsed[0])}");
                    Environment.Exit(1);
                    return;
                }
            }

            // Output header row as-is, sort data rows
            Console.WriteLine(allLines[0]);
            var dataRows = allLines.Skip(1).Zip(parsed.Skip(1), (raw, fields) => (raw, fields));
            var sorted = dataRows.OrderBy(r =>
            {
                string val = colIdx < r.fields.Length ? r.fields[colIdx] : "";
                if (double.TryParse(val, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double num))
                    return (object)num;
                return val;
            }, new MixedComparer());
            var result = descending
                ? (System.Collections.Generic.IEnumerable<(string raw, string[] fields)>)sorted.Reverse()
                : sorted;
            foreach (var row in result)
            {
                Console.WriteLine(row.raw);
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
        Console.WriteLine("  ExcelConsole <file.csv> --sort <col> [--desc]      Headless: sort CSV by column (index or name)");
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

    private static string[] ParseCsvLine(string line)
    {
        var fields = new List<string>();
        var current = new System.Text.StringBuilder();
        bool inQuotes = false;
        for (int i = 0; i < line.Length; i++)
        {
            char ch = line[i];
            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (ch == ',' && !inQuotes)
            {
                fields.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(ch);
            }
        }
        fields.Add(current.ToString());
        return fields.ToArray();
    }

    private class MixedComparer : System.Collections.Generic.IComparer<object>
    {
        public int Compare(object? x, object? y)
        {
            if (x is double dx && y is double dy)
                return dx.CompareTo(dy);
            if (x is double && y is string)
                return -1; // numbers before strings
            if (x is string && y is double)
                return 1;
            return string.Compare(x?.ToString() ?? "", y?.ToString() ?? "",
                StringComparison.OrdinalIgnoreCase);
        }
    }
}
