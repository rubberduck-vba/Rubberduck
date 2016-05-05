using Microsoft.Vbe.Interop;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerAddTestModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly NewUnitTestModuleCommand _newUnitTestModuleCommand;

        public CodeExplorerAddTestModuleCommand(VBE vbe, NewUnitTestModuleCommand newUnitTestModuleCommand)
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