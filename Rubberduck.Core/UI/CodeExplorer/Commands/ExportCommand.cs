using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ExportCommand : CommandBase, IDisposable
    {
        private readonly ISaveFileDialog _saveFileDialog;
        private readonly IProjectsProvider _projectsProvider;
        private readonly Dictionary<ComponentType, string> _exportableFileExtensions = new Dictionary<ComponentType, string>
        {
            { ComponentType.StandardModule, ".bas" },
            { ComponentType.ClassModule, ".cls" },
            { ComponentType.Document, ".cls" },
            { ComponentType.UserForm, ".frm" }
        };

        public ExportCommand(ISaveFileDialog saveFileDialog, IProjectsProvider projectsProvider) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _saveFileDialog = saveFileDialog;
            _saveFileDialog.OverwritePrompt = true;

            _projectsProvider = projectsProvider;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel))
            {
                return false;
            }

            try
            {
                var node = (CodeExplorerComponentViewModel)parameter;
                var componentType = node.Declaration.QualifiedName.QualifiedModuleName.ComponentType;
                return _exportableFileExtensions.Select(s => s.Key).Contains(componentType);
            }
            catch (COMException)
            {
                // thrown when the component reference is stale
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            var node = (CodeExplorerComponentViewModel)parameter;
            var qualifiedModuleName = node.Declaration.QualifiedName.QualifiedModuleName;

            string ext;
            _exportableFileExtensions.TryGetValue(qualifiedModuleName.ComponentType, out ext);

            _saveFileDialog.FileName = qualifiedModuleName.ComponentName + ext;
            var result = _saveFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                var component = _projectsProvider.Component(qualifiedModuleName);
                component.Export(_saveFileDialog.FileName);
            }
        }

        public void Dispose()
        {
            if (_saveFileDialog != null)
            {
                _saveFileDialog.Dispose();
            }
        }
    }
}
