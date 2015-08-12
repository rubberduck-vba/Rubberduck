using Microsoft.Vbe.Interop;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class AddTestModuleCommand : ICommand
    {
        private readonly VBE _vbe;

        public AddTestModuleCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        public void Execute()
        {
            // legacy static class...
            NewUnitTestModuleCommand.NewUnitTestModule(_vbe);
        }
    }
}