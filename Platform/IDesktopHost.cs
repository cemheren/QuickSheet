namespace ExcelConsole.Platform;

/// <summary>
/// Abstraction for embedding the spreadsheet as a desktop background.
/// Implement per-platform: Windows (WorkerW), Linux (X11/Wayland), etc.
/// </summary>
public interface IDesktopHost : IDisposable
{
    void Run(string? csvPath);
}
