using System.Runtime.InteropServices;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test module to the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class AddTestModuleCommand : CommandBase
    {
        private readonly VBE _vbe;

        public AddTestModuleCommand(VBE vbe)
        {
            _vbe = vbe;
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
            // legacy static class...
            NewUnitTestModuleCommand.NewUnitTestModule(_vbe);
        }
    }
}