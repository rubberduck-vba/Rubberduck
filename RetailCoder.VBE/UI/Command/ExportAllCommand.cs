using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Data.Common;
using System.IO;
using System.Windows.Forms;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.CodeExplorer;

namespace Rubberduck.UI.Command
{
    public class ExportAllCommand : CommandBase/*, IDisposable*/
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
            return !_vbe.ActiveVBProject.IsWrappingNullReference && _vbe.ActiveVBProject.Protection != ProjectProtection.Locked; // & _vbe.ActiveVBProject.VBComponents.Count > 0;
        }

        protected override void OnExecute(object parameter)
        {

            var project = _vbe.ActiveVBProject;

            var desc = "Choose a folder to export the source of " + _vbe.ActiveVBProject.Name + " to:";

            // If .GetDirectoryName is passed an empty string for a RootFolder, 
            // it defaults to the Documents library (Win 7+) or equivalent.
            var path = string.Empty;
            if (!string.IsNullOrWhiteSpace(_vbe.ActiveVBProject.FileName))
            {
                path = Path.GetDirectoryName(_vbe.ActiveVBProject.FileName);
            }
            
            using (var _folderBrowser = _factory.CreateFolderBrowser(desc, true, ""))
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
