using Microsoft.Vbe.Interop;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class AddTestMethodExpectedErrorCommand : ICommand
    {
        private readonly VBE _vbe;

        public AddTestMethodExpectedErrorCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        public void Execute()
        {
            // legacy static class...
            NewTestMethodCommand.NewExpectedErrorTestMethod(_vbe);
        }
    }
}