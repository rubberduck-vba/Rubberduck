using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class ExportAllCommand : CommandBase, IDisposable
    {
        private readonly IFolderBrowser _folderBrowserDialog;
        //private readonly Dictionary<ComponentType, string> _exportableFileExtensions = new Dictionary<ComponentType, string>
        //{
        //    { ComponentType.StandardModule, ".bas" },
        //    { ComponentType.ClassModule, ".cls" },
        //    { ComponentType.Document, ".cls" },
        //    { ComponentType.UserForm, ".frm" }
        //};

        public ExportAllCommand(IFolderBrowser folderBrowser) : base(LogManager.GetCurrentClassLogger())
        {
            _folderBrowserDialog = folderBrowser;
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
            return true;
        }

        protected override void OnExecute(object parameter)
        {
            var node = (CodeExplorerComponentViewModel)parameter;
            //var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;
            var project = node.Declaration.Project;

            //string ext;
            //_exportableFileExtensions.TryGetValue(component.Type, out ext);

            //_folderBrowserDialog.FileName = component.Name + ext;
            _folderBrowserDialog.RootFolder = project.FileName;
            var result = _folderBrowserDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                project.ExportSourceFiles(_folderBrowserDialog.SelectedPath);
            }
        }

        public void Dispose()
        {
            if (_folderBrowserDialog != null)
            {
                _folderBrowserDialog.Dispose();
            }
        }
    }
}
