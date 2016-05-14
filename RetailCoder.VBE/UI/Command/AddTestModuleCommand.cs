using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test module to the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class AddTestModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly NewUnitTestModuleCommand _command;

        public AddTestModuleCommand(VBE vbe, NewUnitTestModuleCommand command)
        {
            _vbe = vbe;
            _command = command;
        }

        public override bool CanExecute(object parameter)
        {
            var app = _vbe.HostApplication();
            if (app == null)
            {
                return false;
            }
            else
            {
                // Outlook requires test methods to be located in [ThisOutlookSession] class.
                return app.ApplicationName != "Outlook";
            }
        }

        public override void Execute(object parameter)
        {
            _command.NewUnitTestModule(_vbe.ActiveVBProject);
        }
    }
}
