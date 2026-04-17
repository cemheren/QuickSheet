using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Base form that handles window management for desktop mode: Z-order locking,
/// Alt+Tab hiding, Win+D detection, and WndProc overrides.
/// </summary>
internal class DesktopFormBase : Form
{
    private bool _lockZOrder;
    private IntPtr _winEventHook;
    private NativeMethods.WinEventDelegate? _winEventDelegate;

    // ── Desktop mode ─────────────────────────────────────────────────

    /// <summary>
    /// Configures the form as a desktop replacement: hidden from Alt+Tab.
    /// </summary>
    public void EnterDesktopMode()
    {
        // Hide from Alt+Tab
        var exStyle = (long)NativeMethods.GetWindowLongPtr(Handle, NativeMethods.GWL_EXSTYLE);
        exStyle |= NativeMethods.WS_EX_TOOLWINDOW;
        NativeMethods.SetWindowLongPtr(Handle, NativeMethods.GWL_EXSTYLE, (IntPtr)exStyle);

        // Send to bottom of Z-order so it starts behind all existing windows
        NativeMethods.SetWindowPos(Handle, NativeMethods.HWND_BOTTOM,
            0, 0, 0, 0,
            NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);

        // Lock Z-order so clicking doesn't bring it above other apps
        _lockZOrder = true;

        // Hook foreground window changes to detect Win+D ("Show Desktop").
        // When WorkerW becomes the foreground, Windows has issued "Show Desktop"
        // — make ourselves topmost so we appear above it. When any other window
        // becomes foreground, drop back to non-topmost.
        _winEventDelegate = OnForegroundChanged;
        _winEventHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_SYSTEM_FOREGROUND,
            NativeMethods.EVENT_SYSTEM_FOREGROUND,
            IntPtr.Zero, _winEventDelegate, 0, 0,
            NativeMethods.WINEVENT_OUTOFCONTEXT);

        Invalidate();
    }

    private void OnForegroundChanged(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        var sb = new System.Text.StringBuilder(32);
        NativeMethods.GetClassName(hwnd, sb, sb.Capacity);
        string cls = sb.ToString();

        // Temporarily unlock Z-order so SetWindowPos can change it
        _lockZOrder = false;

        if (cls is "WorkerW" or "Progman")
        {
            // "Show Desktop" was triggered — become topmost to appear above it
            NativeMethods.SetWindowPos(Handle, NativeMethods.HWND_TOPMOST,
                0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
        }
        else
        {
            // A real app took focus — drop topmost so we go behind apps again
            NativeMethods.SetWindowPos(Handle, NativeMethods.HWND_NOTOPMOST,
                0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
        }

        _lockZOrder = true;
    }

    protected override bool ShowWithoutActivation => true;

    protected override void WndProc(ref Message m)
    {
        // Prevent the form from rising in Z-order (covers other apps) while
        // still allowing activation for keyboard focus.
        if (m.Msg == NativeMethods.WM_WINDOWPOSCHANGING && _lockZOrder)
        {
            var pos = Marshal.PtrToStructure<NativeMethods.WINDOWPOS>(m.LParam);
            pos.flags |= NativeMethods.SWP_NOZORDER;
            Marshal.StructureToPtr(pos, m.LParam, false);
        }
        base.WndProc(ref m);
    }

    public virtual void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_winEventHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_winEventHook);
            _winEventHook = IntPtr.Zero;
        }
    }
}
