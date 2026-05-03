#if PLATFORM_WINDOWS
using System.Windows.Forms;
#endif

namespace ExcelConsole.Features
{
    public interface IMode
    {
        void Enter();

        void Exit();

        void Commit();

        bool IsActive();

#if PLATFORM_WINDOWS
        /// <returns>True if handled</returns>
        bool HandleKeyEventWindows(KeyEventArgs e);
#endif

        string GetStatusText();
    }
}
