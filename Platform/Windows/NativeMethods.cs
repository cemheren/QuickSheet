using System.Runtime.InteropServices;

namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Win32 P/Invoke declarations for window management.
/// </summary>
internal static class NativeMethods
{
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
    public static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
    public static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_TOOLWINDOW = 0x00000080;

    public const int WM_WINDOWPOSCHANGING = 0x0046;
    public const int SWP_NOZORDER = 0x0004;

    [StructLayout(LayoutKind.Sequential)]
    public struct WINDOWPOS
    {
        public IntPtr hwnd;
        public IntPtr hwndInsertAfter;
        public int x, y, cx, cy;
        public int flags;
    }
}
