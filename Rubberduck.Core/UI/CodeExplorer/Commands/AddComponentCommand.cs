using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
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

                if (project != null || _vbe.ProjectsCount != 1)
                {
                    return project != null && allowableProjectTypes.Contains(project.Type);
                }

                using (var vbProjects = _vbe.VBProjects)
                using (project = vbProjects[1])
                {                        
                    return project != null && allowableProjectTypes.Contains(project.Type);                        
                }

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

        public void AddComponent(CodeExplorerItemViewModel node, string moduleText)
        {
            var nodeProject = GetDeclaration(node)?.Project;
            if (node != null && nodeProject == null)
            {
                return; //The project is not available.
            }

            string optionCompare;
            using (var hostApp = _vbe.HostApplication())
            {
                optionCompare = hostApp?.ApplicationName == "Access" ? "Option Compare Database" :
                    string.Empty;
            }

            using (var components = node != null
                ? nodeProject.VBComponents
                : ComponentsCollectionFromActiveProject())
            {
                var folderAnnotation = $"'@Folder(\"{GetFolder(node)}\")";
                var fileName = CreateTempTextFile(moduleText);

                using (var newComponent = components.Import(fileName))
                {
                    using (var codeModule = newComponent.CodeModule)
                    {
                        if (optionCompare.Length > 0)
                        {
                            codeModule.InsertLines(1, optionCompare);
                        }
                        if (folderAnnotation.Length > 0)
                        {
                            codeModule.InsertLines(1, folderAnnotation);
                        }
                        codeModule.CodePane.Show();
                    }
                }
                File.Delete(fileName);
            }
        }

        private static string CreateTempTextFile(string moduleText)
        {
            var tempFolder = ApplicationConstants.RUBBERDUCK_TEMP_PATH;
            if (!Directory.Exists(tempFolder))
            {
                Directory.CreateDirectory(tempFolder);
            }
            var filePath = Path.Combine(tempFolder, Path.GetRandomFileName());
            File.WriteAllText(filePath, moduleText);
            return filePath;
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