#if PLATFORM_WINDOWS
using System.Windows.Forms;
#elif PLATFORM_LINUX
using static ExcelConsole.Platform.Linux.X11Methods;
#endif

namespace ExcelConsole.Features
{
    internal interface IMode
    {
        void Enter();

        void Exit();

        void Commit();

        bool IsActive();

#if PLATFORM_WINDOWS
        /// <returns>True if handled</returns>
        bool HandleKeyEventWindows(KeyEventArgs e);
#elif PLATFORM_LINUX
        /// <returns>True if handled</returns>
        bool HandleKeyEventLinux(ulong keysym, ref XKeyEvent keyEvent, bool ctrl);
#endif

        string GetStatusText();
    }
}
