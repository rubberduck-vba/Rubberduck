using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class AddClassModuleCommand : CommandBase
    {
        private readonly VBE _vbe;

        public AddClassModuleCommand(VBE vbe) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            try
            {
                return GetDeclaration(parameter) != null || _vbe.VBProjects.Count == 1;
            }
            catch (COMException)
            {
                return false;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            if (parameter != null)
            {
                GetDeclaration(parameter).Project.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
            }
            else
            {
                _vbe.VBProjects.Item(1).VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
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
