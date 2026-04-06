using System.Runtime.InteropServices;
using ExcelConsole;

public class Program
{
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [STAThread]
    public static void Main(string[] args)
    {
        string? csvPath = args.FirstOrDefault(a => !a.StartsWith("--"));
        bool desktopMode = args.Contains("--desktop");

        if (desktopMode)
        {
            IntPtr console = GetConsoleWindow();
            if (console != IntPtr.Zero)
                ShowWindow(console, 0); // SW_HIDE
            if (!OperatingSystem.IsWindows())
            {
                Console.Error.WriteLine("Desktop mode is currently only supported on Windows.");
                return;
            }

            using var host = new ExcelConsole.Platform.Windows.WindowsDesktopHost();
            host.Run(csvPath);
        }
        else
        {
            var app = new SpreadsheetApp(csvPath);
            app.Run();
        }
    }
}
