using Microsoft.Vbe.Interop;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_AddStdModuleCommand : CommandBase
    {
        public override bool CanExecute(object parameter)
        {
            return GetDeclaration(parameter) != null;
        }

        public override void Execute(object parameter)
        {
            GetDeclaration(parameter).Project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
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
