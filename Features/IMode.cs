using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConsole.Features
{
    public interface IMode
    {
        void Enter();

        void Exit();

        void Commit();

        bool IsActive();

        /// <returns>True if handled</returns>
        bool HandleKeyEvent(KeyEventArgs e);

        string GetStatusText();
    }
}
