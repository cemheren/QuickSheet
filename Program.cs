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
            grid.SaveToMarkdown(outPath);
            Console.WriteLine($"Wrote markdown: {outPath}");
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

    private static void PrintHelp()
    {
        Console.WriteLine("QuickSheet — interactive terminal spreadsheet + desktop wallpaper");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  ExcelConsole [<file.csv>]                  TUI mode (default)");
        Console.WriteLine("  ExcelConsole [<file.csv>] --desktop        Embed as desktop wallpaper");
        Console.WriteLine("  ExcelConsole <file.csv> --export-md <out.md>  Headless: CSV → Markdown table");
        Console.WriteLine("  ExcelConsole --help                        Show this help");
        Console.WriteLine();
        Console.WriteLine("Cell prefixes (TUI / desktop modes):");
        Console.WriteLine("  r: <cmd>          Runnable command. Press Enter to launch.");
        Console.WriteLine("  i: <cmd>          Inline subprocess. Output streams back into the cell.");
        Console.WriteLine("  s: 1,2,3,...      Sparkline (unicode block bars). Also accepts range: s: A1::A10");
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
