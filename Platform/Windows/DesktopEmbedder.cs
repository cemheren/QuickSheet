namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Finds the WorkerW window behind desktop icons where we can embed our form.
/// Uses the standard Progman 0x052C message trick used by live wallpaper engines.
/// </summary>
internal static class DesktopEmbedder
{
    /// <summary>
    /// Asks Windows to spawn a WorkerW behind the desktop icons, then returns its handle.
    /// </summary>
    public static IntPtr GetDesktopWorkerW()
    {
        IntPtr progman = NativeMethods.FindWindow("Progman", null);
        if (progman == IntPtr.Zero)
            return IntPtr.Zero;

        // Tell Progman to spawn a WorkerW window behind the desktop icons
        NativeMethods.SendMessage(progman, 0x052C, new IntPtr(0xD), new IntPtr(0x1));

        // Give Windows a moment to create the WorkerW
        Thread.Sleep(100);

        IntPtr workerW = IntPtr.Zero;
        NativeMethods.EnumWindows((hwnd, _) =>
        {
            var shell = NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
            if (shell != IntPtr.Zero)
            {
                // The WorkerW we want is the next sibling after the one containing SHELLDLL_DefView
                workerW = NativeMethods.FindWindowEx(IntPtr.Zero, hwnd, "WorkerW", null);
            }
            return true;
        }, IntPtr.Zero);

        return workerW;
    }
}
