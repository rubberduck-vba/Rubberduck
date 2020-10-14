using System.IO;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    public class ExportAllCommand : ComCommandBase 
    {
        private readonly IVBE _vbe;
        private readonly IFileSystemBrowserFactory _factory;
        private readonly IProjectsProvider _projectsProvider;

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
            var desc = string.Format(RubberduckUI.ExportAllCommand_SaveAsDialog_Title, project.Name);

            // If .GetDirectoryName is passed an empty string for a RootFolder, 
            // it defaults to the Documents library (Win 7+) or equivalent.
            var path = string.IsNullOrWhiteSpace(project.FileName)
                ? string.Empty
                : Path.GetDirectoryName(project.FileName);

            using (var _folderBrowser = _factory.CreateFolderBrowser(desc, true, path))
            {
                var result = _folderBrowser.ShowDialog();

                if (result == DialogResult.OK)
                {
                    project.ExportSourceFiles(_folderBrowser.SelectedPath);
                }
            }
        }
    }
}