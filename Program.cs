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
        bool desktopMode = args.Contains("--desktop");

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

        if (desktopMode)
        {
#if PLATFORM_WINDOWS
            HideConsoleWindow();
            using var host = new ExcelConsole.Platform.Windows.WindowsDesktopHost();
            host.Run(csvPath);
#elif PLATFORM_LINUX
            string sessionType = Environment.GetEnvironmentVariable("XDG_SESSION_TYPE") ?? "";
            if (sessionType.Equals("wayland", StringComparison.OrdinalIgnoreCase))
            {
                Console.Error.WriteLine("Warning: Desktop mode requires an X11 session.");
                Console.Error.WriteLine("You are running Wayland. Select 'GNOME on Xorg' at the login screen,");
                Console.Error.WriteLine("or set DISPLAY and try with XWayland (may not work as true desktop layer).");
                Console.Error.WriteLine();
                Console.Error.WriteLine("Attempting via XWayland anyway...");
            }

            using var host = new ExcelConsole.Platform.Linux.LinuxDesktopHost();
            host.Run(csvPath);
#else
            Console.Error.WriteLine("Desktop mode is only supported on Windows and Linux (X11).");
#endif
        }
        else
        {
            var app = new SpreadsheetApp(csvPath);
            app.Run();
        }
    }

    private const string Version = "0.3.0";

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
        Console.WriteLine("QuickSheet — interactive terminal spreadsheet + desktop wallpaper");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  ExcelConsole [<file.csv>]                  TUI mode (default)");
        Console.WriteLine("  ExcelConsole [<file.csv>] --desktop        Embed as desktop wallpaper");
        Console.WriteLine("  ExcelConsole <file.csv> --export-md <out.md>  Headless: CSV → Markdown table (use - for stdout)");
        Console.WriteLine("  ExcelConsole --help                        Show this help");
        Console.WriteLine("  ExcelConsole --version                     Show version");
        Console.WriteLine("  ExcelConsole --list-extensions             List installed extensions");
        Console.WriteLine();
        Console.WriteLine("Cell prefixes (TUI / desktop modes):");
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
}
