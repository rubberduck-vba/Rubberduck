using Microsoft.Vbe.Interop;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_AddTestModuleCommand : CommandBase
    {
        private readonly NewUnitTestModuleCommand _newUnitTestModuleCommand;

        public CodeExplorer_AddTestModuleCommand(NewUnitTestModuleCommand newUnitTestModuleCommand)
        {
            _newUnitTestModuleCommand = newUnitTestModuleCommand;
        }

        public override bool CanExecute(object parameter)
        {
            return GetDeclaration(parameter) != null;
        }

        public override void Execute(object parameter)
        {
            _newUnitTestModuleCommand.NewUnitTestModule(GetDeclaration(parameter).Project);
        }

        private Declaration GetDeclaration(object parameter)
        {
            var node = parameter as CodeExplorerItemViewModel;
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return node == null ? null : ((ICodeExplorerDeclarationViewModel)node).Declaration;
        }
    }
}