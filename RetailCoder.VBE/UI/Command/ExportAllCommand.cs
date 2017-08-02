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
        //private IFolderBrowser _folderBrowser;
        private readonly string _filePath;

        public ExportAllCommand(IVBE vbe, IFolderBrowserFactory folderBrowserFactory) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _factory = folderBrowserFactory;
            //_folderBrowser = folderBrowserFactory.CreateFolderBrowser("Select a directory to Export Project as Source Files...", true);
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            //if (!(parameter is CodeExplorerComponentViewModel))
            //{
            //    return false;
            //}

            //try
            //{
            //    var node = (CodeExplorerComponentViewModel)parameter;
            //    var componentType = node.Declaration.QualifiedName.QualifiedModuleName.ComponentType;
            //    return _exportableFileExtensions.Select(s => s.Key).Contains(componentType);
            //}
            //catch (COMException)
            //{
            //    // thrown when the component reference is stale
            //    return false;
            //}
            return !_vbe.ActiveVBProject.IsWrappingNullReference && _vbe.ActiveVBProject.Protection != ProjectProtection.Locked;
        }

        protected override void OnExecute(object parameter)
        {

            var project = _vbe.ActiveVBProject;
            //var project = (IVBProject)parameter;

            using (var _folderBrowser = _factory.CreateFolderBrowser("Select a folder"))
            {
                _folderBrowser.ShowNewFolderButton = false;  // test to see if the dialog changes when shown
                _folderBrowser.Description = "asdf";  // test to see if the dialog changes when shown
                _folderBrowser.RootFolder = Path.GetDirectoryName(_vbe.ActiveVBProject.FileName);

                var result = _folderBrowser.ShowDialog();

                if (result == DialogResult.OK)
                {
                    project.ExportSourceFiles(_folderBrowser.SelectedPath);
                }
            }


            //_folderBrowserDialog.FileName = component.Name + ext;
            //_folderBrowser.RootFolder = project.FileName;


        }

        //public void Dispose()
        //{
        //    if (_folderBrowser != null)
        //    {
        //        _folderBrowser.Dispose();
        //    }
        //}
    }
}
