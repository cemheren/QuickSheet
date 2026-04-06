using ExcelConsole;

string? csvPath = args.FirstOrDefault(a => !a.StartsWith("--"));
bool desktopMode = args.Contains("--desktop");

if (desktopMode)
{
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
