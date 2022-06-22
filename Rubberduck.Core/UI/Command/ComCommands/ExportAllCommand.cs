using Path = System.IO.Path;
using Directory = System.IO.Directory;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;

namespace Rubberduck.UI.Command.ComCommands
{
    public class ExportAllCommand : ComCommandBase 
    {
        private readonly IVBE _vbe;
        private readonly IProjectsProvider _projectsProvider;

        private readonly IFileSystemBrowserFactory _factory;

        private static Dictionary<string,string> _projectExportFolderpaths;

        public ExportAllCommand(
            IVBE vbe, 
            IFileSystemBrowserFactory folderBrowserFactory, 
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _factory = folderBrowserFactory;
            _projectsProvider = projectsProvider;
            if (_projectExportFolderpaths is null)
            {
                _projectExportFolderpaths = new Dictionary<string, string>();
            }

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (_vbe.Kind == VBEKind.Standalone)
            {
                return false;
            }

            if (!(parameter is CodeExplorerProjectViewModel) 
                && parameter is CodeExplorerItemViewModel)
            {
                return false;
            }

            switch (parameter)
            {
                case CodeExplorerProjectViewModel projectNode:
                    var nodeProject = projectNode.Declaration != null
                        ? _projectsProvider.Project(projectNode.Declaration.ProjectId)
                        : null;
                    return Evaluate(nodeProject);
                case IVBProject project:
                    return Evaluate(project);
            }

            using (var activeProject = _vbe.ActiveVBProject)
            {
                return Evaluate(activeProject);
            }
        }

        private bool Evaluate(IVBProject project)
        {
            if (project == null || project.IsWrappingNullReference)
            {
                return false;
            }

            using (var compontents = project.VBComponents)
            {
                return compontents.Count > 0;
            }

        }

        protected override void OnExecute(object parameter)
        {
            switch (parameter)
            {
                case CodeExplorerProjectViewModel projectNode:
                    var nodeProject = projectNode.Declaration != null
                        ? _projectsProvider.Project(projectNode.Declaration.ProjectId)
                        : null;
                    if (nodeProject == null)
                    {
                        return;
                    }
                    Export(nodeProject);
                    break;
                case IVBProject vbproject:
                    Export(vbproject);
                    break;
                default:
                {
                    using (var project = _vbe.ActiveVBProject)
                    {
                        Export(project);
                    }
                    break;
                }
            }
        }

        private void Export(IVBProject project)
        {
            var initialFolderBrowserPath = GetInitialFolderBrowserPath(project);

            var desc = string.Format(RubberduckUI.ExportAllCommand_SaveAsDialog_Title, project.Name);

            using (var _folderBrowser = _factory.CreateFolderBrowser(desc, true, initialFolderBrowserPath))
            {
                var result = _folderBrowser.ShowDialog();

                if (result == DialogResult.OK)
                {
                    _projectExportFolderpaths[project.FileName] = _folderBrowser.SelectedPath;
                    project.ExportSourceFiles(_folderBrowser.SelectedPath);
                }
            }
        }

        //protected scope to support testing
        protected string GetInitialFolderBrowserPath(IVBProject project)
        {
            if (!(string.IsNullOrWhiteSpace(project.FileName))
                && _projectExportFolderpaths.TryGetValue(project.FileName, out string initialFolderBrowserPath))
            {
                if (FolderExists(initialFolderBrowserPath))
                {
                    //Return the cached folderpath of the previous ExportAllCommand process
                    return _projectExportFolderpaths[project.FileName];
                }

                //The folder used in the previous ExportAllComand process no longer exists, remove the cached folderpath
                _projectExportFolderpaths.Remove(project.FileName);
            }

            //The folder of the workbook, or an empty string
            initialFolderBrowserPath = GetDefaultExportFolder(project.FileName);

            if (!string.IsNullOrEmpty(initialFolderBrowserPath))
            {
                _projectExportFolderpaths.Add(project.FileName, initialFolderBrowserPath);
            }

            return initialFolderBrowserPath;
        }

        //protected scope to support testing
        protected string GetDefaultExportFolder(string projectFileName)
        {
            // If .GetDirectoryName is passed an empty string for a RootFolder, 
            // it defaults to the Documents library (Win 7+) or equivalent.
            return string.IsNullOrWhiteSpace(projectFileName)
                ? string.Empty
                : Path.GetDirectoryName(projectFileName);
        }

        //protected virtual to support testing
        protected virtual bool FolderExists(string path)
        {
            return Directory.Exists(path);
        }
    }
}