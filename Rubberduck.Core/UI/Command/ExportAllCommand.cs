using System.IO;
using System.Windows.Forms;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.Command
{
    public class ExportAllCommand : CommandBase 
    {
        private readonly IVBE _vbe;
        private readonly IFolderBrowserFactory _factory;

        public ExportAllCommand(IVBE vbe, IFolderBrowserFactory folderBrowserFactory) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _factory = folderBrowserFactory;
        }

        protected override bool EvaluateCanExecute(object parameter)
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
            var projectNode = parameter as CodeExplorerProjectViewModel;

            var vbproject = parameter as IVBProject;

            var project = projectNode?.Declaration.Project ?? vbproject ?? _vbe.ActiveVBProject;
            
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