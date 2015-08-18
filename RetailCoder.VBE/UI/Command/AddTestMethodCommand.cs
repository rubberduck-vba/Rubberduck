using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodCommand : CommandBase
    {
        private readonly VBE _vbe;

        public AddTestMethodCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        public override void Execute(object parameter)
        {
            // legacy static class...
            NewTestMethodCommand.NewTestMethod(_vbe);
        }
    }
}