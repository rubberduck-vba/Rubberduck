using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddComponentCommand
    {
        private readonly IVBE _vbe;

        public AddComponentCommand(IVBE vbe)
        {
            _vbe = vbe;
        }

        public bool CanAddComponent(CodeExplorerItemViewModel parameter, IEnumerable<ProjectType> allowableProjectTypes)
        {
            try
            {
                var project = GetDeclaration(parameter)?.Project;

                if (project == null && _vbe.ProjectsCount == 1)
                {
                    using (var vbProjects = _vbe.VBProjects)
                    using (project = vbProjects[1])
                    {                        
                        return project != null && allowableProjectTypes.Contains(project.Type);                        
                    }
                }

                return project != null && allowableProjectTypes.Contains(project.Type);

            }
            catch (COMException)
            {
                return false;
            }
        }

        public void AddComponent(CodeExplorerItemViewModel node, ComponentType type)
        {
            var nodeProject = GetDeclaration(node)?.Project;
            if (node != null && nodeProject == null)
            {
                return; //The project is not available.
            }

            using (var components = node != null
                ? nodeProject.VBComponents
                : ComponentsCollectionFromActiveProject())
            {
                var folderAnnotation = $"'@Folder(\"{GetFolder(node)}\")";

                using (var newComponent = components.Add(type))
                {
                    using (var codeModule = newComponent.CodeModule)
                    {
                        codeModule.InsertLines(1, folderAnnotation);
                    }
                }
            }
        }

        private IVBComponents ComponentsCollectionFromActiveProject()
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                return activeProject.VBComponents;
            }
        }

        private Declaration GetDeclaration(CodeExplorerItemViewModel node)
        {
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return (node as ICodeExplorerDeclarationViewModel)?.Declaration;
        }
        private string GetActiveProjectName()
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                return activeProject.Name;
            }
        }
        private string GetFolder(CodeExplorerItemViewModel node)
        {
            switch (node)
            {
                case null:
                    return GetActiveProjectName();
                case ICodeExplorerDeclarationViewModel declarationNode:
                    return string.IsNullOrEmpty(declarationNode.Declaration.CustomFolder)
                        ? GetActiveProjectName()
                        : declarationNode.Declaration.CustomFolder.Replace("\"", string.Empty);
                default:
                    return ((CodeExplorerCustomFolderViewModel)node).FullPath;
            }
        }
    }
}