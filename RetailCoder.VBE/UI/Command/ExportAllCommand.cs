using System.IO;
using System.Windows.Forms;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
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
            return !_vbe.ActiveVBProject.IsWrappingNullReference && _vbe.ActiveVBProject.VBComponents.Count > 0;
        }

        protected override void OnExecute(object parameter)
        {
            IVBProject project;
            if (parameter == null)
            {
                project = _vbe.ActiveVBProject;
            }
            else
            {
                project = (IVBProject)parameter;
            }

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
