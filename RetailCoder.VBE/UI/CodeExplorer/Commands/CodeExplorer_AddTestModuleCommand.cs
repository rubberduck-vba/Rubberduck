using Microsoft.Vbe.Interop;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_AddTestModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly NewUnitTestModuleCommand _newUnitTestModuleCommand;

        public CodeExplorer_AddTestModuleCommand(VBE vbe, NewUnitTestModuleCommand newUnitTestModuleCommand)
        {
            _vbe = vbe;
            _newUnitTestModuleCommand = newUnitTestModuleCommand;
        }

        public override bool CanExecute(object parameter)
        {
            return _vbe.ActiveVBProject != null;
        }

        public override void Execute(object parameter)
        {
            _newUnitTestModuleCommand.NewUnitTestModule();
        }
    }
}