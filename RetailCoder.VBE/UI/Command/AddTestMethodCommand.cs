using Microsoft.Vbe.Interop;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class AddTestMethodCommand : ICommand
    {
        private readonly VBE _vbe;

        public AddTestMethodCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        public void Execute()
        {
            // legacy static class...
            NewTestMethodCommand.NewTestMethod(_vbe);
        }
    }
}