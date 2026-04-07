using static ExcelConsole.Platform.Linux.X11Methods;

namespace ExcelConsole.Platform.Linux;

/// <summary>
/// Linux implementation of IDesktopHost.
/// Opens an X11 display, creates a DesktopWindow with _NET_WM_WINDOW_TYPE_DESKTOP,
/// and runs the X11 event loop. Requires an X11 session (not Wayland).
/// </summary>
public class LinuxDesktopHost : Platform.IDesktopHost
{
    private IntPtr _display;
    private DesktopWindow? _window;

    public void Run(string? csvPath)
    {
        _display = XOpenDisplay(null);
        if (_display == IntPtr.Zero)
            return;

        try
        {
            _window = new DesktopWindow(_display, csvPath);
            _window.Run();
        }
        finally
        {
            _window?.Dispose();
            if (_display != IntPtr.Zero)
                XCloseDisplay(_display);
        }
    }

    public void Dispose()
    {
        _window?.Stop();
        _window?.Dispose();
        _window = null;
    }
}
