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
            // Outlook requires test methods to be located in [ThisOutlookSession] class.
            return _vbe.HostApplication().ApplicationName != "Outlook";
        }

        public override void Execute(object parameter)
        {
            // legacy static class...
            NewUnitTestModuleCommand.NewUnitTestModule(_vbe);
        }
    }
}