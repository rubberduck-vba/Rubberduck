using System.IO;
using System.Windows.Forms;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.CodeExplorer.Commands;

namespace Rubberduck.UI.Command
{
    [CodeExplorerCommand]
    public class ExportAllCommand : CommandBase 
    {
        private readonly IVBE _vbe;
        private IFolderBrowserFactory _factory;

        public ExportAllCommand(IVBE vbe, IFolderBrowserFactory folderBrowserFactory) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _factory = folderBrowserFactory;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (parameter == null)
            {
                return Evaluate(_vbe.ActiveVBProject);
            }
            else if (parameter is CodeExplorerProjectViewModel)
            {
                var node = (CodeExplorerProjectViewModel)parameter;
                return Evaluate((IVBProject)node.Declaration.Project);
            }
            else if (parameter is IVBProject)
            {
                return Evaluate((IVBProject)parameter);
            }
            else { return false; }
        }

        private bool Evaluate(IVBProject project)
        {
            return !project.IsWrappingNullReference && project.VBComponents.Count > 0;
        }

        protected override void OnExecute(object parameter)
        {
            IVBProject project;
            if (parameter == null) { project = (_vbe.ActiveVBProject); }
            else if (parameter is CodeExplorerProjectViewModel)
            {
                CodeExplorerProjectViewModel projectVM = (CodeExplorerProjectViewModel)parameter;
                project = projectVM.Declaration.Project;
            }
            else { project = (IVBProject)parameter; } // for unit test in ExportAllCommand.cs

            var desc = "Choose a folder to export the source of " + project.Name + " to:";

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