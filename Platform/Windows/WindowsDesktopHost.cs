using System.Windows.Forms;

namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Windows implementation of IDesktopHost.
/// Launches a WinForms application and embeds it into the desktop WorkerW layer.
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

        // Force window handle creation without showing, so we can embed first
        _ = _form.Handle;
        _form.EmbedIntoDesktop();

        // Application.Run makes the form visible — it's already parented to WorkerW
        Application.Run(_form);
    }

    public void Dispose()
    {
        _form?.Dispose();
    }
}
