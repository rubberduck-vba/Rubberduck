using Microsoft.Vbe.Interop;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_AddUserFormCommand : CommandBase
    {
        private readonly VBE _vbe;

        public CodeExplorer_AddUserFormCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        public override bool CanExecute(object parameter)
        {
            return GetDeclaration(parameter) != null || _vbe.VBProjects.Count == 1;
        }

        public override void Execute(object parameter)
        {
            if (parameter != null)
            {
                GetDeclaration(parameter).Project.VBComponents.Add(vbext_ComponentType.vbext_ct_MSForm);
            }
            else
            {
                _vbe.VBProjects.Item(1).VBComponents.Add(vbext_ComponentType.vbext_ct_MSForm);
            }
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
