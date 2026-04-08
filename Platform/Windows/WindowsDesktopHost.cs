using System.Windows.Forms;

namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Windows implementation of IDesktopHost.
/// Launches a borderless WinForms window that acts as a desktop replacement.
/// </summary>
public class WindowsDesktopHost : Platform.IDesktopHost
{
    private DesktopForm? _form;

    public void Run(string? csvPath)
    {
        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        _form = new DesktopForm(csvPath);
        _form.Shown += (_, _) => _form.EnterDesktopMode();

        Application.Run(_form);
    }

    public void Dispose()
    {
        _form?.Dispose();
    }
}
