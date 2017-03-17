using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class AddClassModuleCommand : CommandBase
    {
        private readonly IVBE _vbe;

        public AddClassModuleCommand(IVBE vbe) : base(LogManager.GetCurrentClassLogger())
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
            var folderAnnotation = $"'@Folder(\"{GetFolder(parameter)}\")";

            if (parameter != null)
            {
                var components = GetDeclaration(parameter).Project.VBComponents;
                var newComponent = components.Add(ComponentType.ClassModule);
                newComponent.CodeModule.AddFromString(folderAnnotation);
            }
            else
            {
                var project = _vbe.ActiveVBProject;
                var components = project.VBComponents;
                var newComponent = components.Add(ComponentType.ClassModule);
                newComponent.CodeModule.AddFromString(folderAnnotation);
            }
        }

        private Declaration GetDeclaration(object parameter)
        {
            var node = parameter as CodeExplorerItemViewModel;
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return ((ICodeExplorerDeclarationViewModel) node)?.Declaration;
        }

        private string GetFolder(object parameter)
        {
            if (parameter == null)
            {
                return "VBAProject";
            }

            var declarationNode = parameter as ICodeExplorerDeclarationViewModel;
            if (declarationNode != null)
            {
                return string.IsNullOrEmpty(declarationNode.Declaration.CustomFolder)
                    ? "VBAProject"
                    : declarationNode.Declaration.CustomFolder.Replace("\"", string.Empty);
            }

            return ((CodeExplorerCustomFolderViewModel)parameter).FullPath;
        }
    }
}
