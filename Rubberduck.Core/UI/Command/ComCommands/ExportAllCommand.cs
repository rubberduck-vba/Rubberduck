using System.IO;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    public class ExportAllCommand : ComCommandBase 
    {
        private readonly IVBE _vbe;
        private readonly IFileSystemBrowserFactory _factory;

        public ExportAllCommand(
            IVBE vbe, 
            IFileSystemBrowserFactory folderBrowserFactory, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _factory = folderBrowserFactory;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (_vbe.Kind == VBEKind.Standalone)
            {
                return false;
            }

            if (!(parameter is CodeExplorerProjectViewModel) && parameter is CodeExplorerItemViewModel)
            {
                return false;
            }

            switch (parameter)
            {
                case CodeExplorerProjectViewModel projectNode:
                    return Evaluate(projectNode.Declaration.Project);
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
                case CodeExplorerProjectViewModel projectNode when projectNode.Declaration.Project != null:
                    Export(projectNode.Declaration.Project);
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
            var path = string.Empty;
            if (!string.IsNullOrWhiteSpace(project.FileName))
            {
                path = Path.GetDirectoryName(project.FileName);
            }

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