using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.DisposableWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class AddStdModuleCommand : CommandBase
    {
        private readonly VBE _vbe;

        public AddStdModuleCommand(VBE vbe) : base(LogManager.GetCurrentClassLogger())
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
                // could be that _vbe.VBProjects reference is stale?
                return false;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            if (parameter != null)
            {
                using (var components = GetDeclaration(parameter).Project.VBComponents)
                {
                    components.Add(ComponentType.StandardModule);
                }
            }
            else
            {
                using (var project = _vbe.ActiveVBProject)
                using (var components = project.VBComponents)
                {
                    components.Add(ComponentType.StandardModule);
                }
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
