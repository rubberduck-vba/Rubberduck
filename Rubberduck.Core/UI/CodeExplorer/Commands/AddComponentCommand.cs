using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public abstract class AddComponentCommandBase : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        protected AddComponentCommandBase(IVBE vbe)
        {
            Vbe = vbe;
        }

        protected IVBE Vbe { get; }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        public abstract IEnumerable<ProjectType> AllowableProjectTypes { get; }

        public abstract ComponentType ComponentType { get; }

        protected override void OnExecute(object parameter)
        {
            AddComponent(parameter as CodeExplorerItemViewModel);
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!base.EvaluateCanExecute(parameter) || 
                !(parameter is CodeExplorerItemViewModel node) ||
                !(node.Declaration?.Project is IVBProject project))
            {
                return false;
            }

            try
            {
                if (Vbe.ProjectsCount != 1)
                {
                    return AllowableProjectTypes.Contains(project.Type);
                }

                using (var vbProjects = Vbe.VBProjects)
                using (project = vbProjects[1])
                {
                    return AllowableProjectTypes.Contains(project.Type);
                }

            }
            catch (COMException)
            {
                return false;
            }
        }

        protected void AddComponent(CodeExplorerItemViewModel node)
        {
            var nodeProject = node?.Declaration.Project;
            if (node != null && nodeProject == null)
            {
                return; //The project is not available.
            }

            using (var components = node != null
                ? nodeProject.VBComponents
                : ComponentsCollectionFromActiveProject())
            {
                var folderAnnotation = (node is CodeExplorerCustomFolderViewModel folder) ? folder.FolderAttribute : $"'@Folder(\"{GetFolder(node)}\")";

                using (var newComponent = components.Add(ComponentType))
                {
                    using (var codeModule = newComponent.CodeModule)
                    {
                        codeModule.InsertLines(1, folderAnnotation);
                    }
                }
            }
        }

        protected void AddComponent(CodeExplorerItemViewModel node, string moduleText)
        {
            var nodeProject = node?.Declaration?.Project;
            if (node != null && nodeProject == null)
            {
                return; //The project is not available.
            }

            string optionCompare;
            using (var hostApp = Vbe.HostApplication())
            {
                optionCompare = hostApp?.ApplicationName == "Access" ? "Option Compare Database" :
                    string.Empty;
            }

            using (var components = node != null
                ? nodeProject.VBComponents
                : ComponentsCollectionFromActiveProject())
            {
                var folderAnnotation = (node is CodeExplorerCustomFolderViewModel folder) ? folder.FolderAttribute : $"'@Folder(\"{GetFolder(node)}\")";
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
            using (var activeProject = Vbe.ActiveVBProject)
            {
                return activeProject.VBComponents;
            }
        }

        private string GetActiveProjectName()
        {
            using (var activeProject = Vbe.ActiveVBProject)
            {
                return activeProject.Name;
            }
        }

        private string GetFolder(CodeExplorerItemViewModel node)
        {
            return string.IsNullOrEmpty(node?.Declaration?.CustomFolder)
                ? GetActiveProjectName()
                : node.Declaration.CustomFolder.Replace("\"", string.Empty);
        }
    }
}