using System.Runtime.InteropServices;
using ExcelConsole;

public class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        string? csvPath = args.FirstOrDefault(a => !a.StartsWith("--"));
        bool desktopMode = args.Contains("--desktop");

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
