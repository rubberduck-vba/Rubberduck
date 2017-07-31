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
        private readonly IFolderBrowser _folderBrowser;
        //private readonly Dictionary<ComponentType, string> _exportableFileExtensions = new Dictionary<ComponentType, string>
        //{
        //    { ComponentType.StandardModule, ".bas" },
        //    { ComponentType.ClassModule, ".cls" },
        //    { ComponentType.Document, ".cls" },
        //    { ComponentType.UserForm, ".frm" }
        //};

        public ExportAllCommand(IFolderBrowserFactory folderBrowserFactory) : base(LogManager.GetCurrentClassLogger())
        {
            _folderBrowser = folderBrowserFactory.CreateFolderBrowser("description", true, @"c:\");
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
