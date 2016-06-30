using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_AddStdModuleCommand : CommandBase
    {
        private readonly VBE _vbe;

        public CodeExplorer_AddStdModuleCommand(VBE vbe) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return GetDeclaration(parameter) != null || _vbe.VBProjects.Count == 1;
        }

        protected override void ExecuteImpl(object parameter)
        {
            if (parameter != null)
            {
                GetDeclaration(parameter).Project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            }
            else
            {
                _vbe.VBProjects.Item(1).VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
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
