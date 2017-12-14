using System.IO;
using System.Windows.Forms;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    [CodeExplorerCommand]
    public class ExportAllCommand : CommandBase 
    {
        private readonly IVBE _vbe;
        private readonly IFolderBrowserFactory _factory;

        public ExportAllCommand(IVBE vbe, IFolderBrowserFactory folderBrowserFactory) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _factory = folderBrowserFactory;
        }

        public override RubberduckHotkey Hotkey => RubberduckHotkey.ExportActiveProject;

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerProjectViewModel) && parameter is CodeExplorerItemViewModel)
            {
                return false;
            }

            var projectNode = parameter as CodeExplorerProjectViewModel;

            var project = parameter as IVBProject;

            return Evaluate(projectNode?.Declaration.Project ?? project ?? _vbe.ActiveVBProject);

        }

        private bool Evaluate(IVBProject project)
        {
            return project != null && !project.IsWrappingNullReference && project.VBComponents.Count > 0;
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