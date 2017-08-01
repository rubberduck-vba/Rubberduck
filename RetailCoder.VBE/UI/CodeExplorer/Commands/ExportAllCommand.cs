using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
    public class ExportAllCommand : CommandBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly IFolderBrowser _folderBrowser;
        private readonly string _filePath;
        //private readonly Dictionary<ComponentType, string> _exportableFileExtensions = new Dictionary<ComponentType, string>
        //{
        //    { ComponentType.StandardModule, ".bas" },
        //    { ComponentType.ClassModule, ".cls" },
        //    { ComponentType.Document, ".cls" },
        //    { ComponentType.UserForm, ".frm" }
        //};

        public ExportAllCommand(IVBE vbe, IFolderBrowser folderBrowser) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _filePath = vbe.ActiveVBProject.FileName;
            _folderBrowser = folderBrowser.CreateFolderBrowser("Select a directory to Export Project as Source Files...", true, @"c:\");
            //_folderBrowserDialog.OverwritePrompt = true;
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
            //var node = (CodeExplorerComponentViewModel)parameter;
            //var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;
            //var project = _vbe.ActiveVBProject;
            var project = (IVBProject)parameter;

            //string ext;
            //_exportableFileExtensions.TryGetValue(component.Type, out ext);

            //_folderBrowserDialog.FileName = component.Name + ext;
            //_folderBrowser.RootFolder = project.FileName;
            _folderBrowser.RootFolder = @"C:\";
            var result = _folderBrowser.ShowDialog();

            if (result == DialogResult.OK)
            {
                project.ExportSourceFiles(_folderBrowser.SelectedPath);
            }
        }

        public void Dispose()
        {
            if (_folderBrowser != null)
            {
                _folderBrowser.Dispose();
            }
        }
    }
}
