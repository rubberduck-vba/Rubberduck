using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class RemoveCommand : CommandBase, IDisposable
    {
        private readonly ISaveFileDialog _saveFileDialog;
        private readonly IMessageBox _messageBox;
        private readonly IProjectsProvider _projectsProvider;

        private readonly Dictionary<ComponentType, string> _exportableFileExtensions = new Dictionary<ComponentType, string>
        {
            { ComponentType.StandardModule, ".bas" },
            { ComponentType.ClassModule, ".cls" },
            { ComponentType.Document, ".cls" },
            { ComponentType.UserForm, ".frm" }
        };

        public RemoveCommand(ISaveFileDialog saveFileDialog, IMessageBox messageBox, IProjectsProvider projectsProvider) : base(LogManager.GetCurrentClassLogger())
        {
            _saveFileDialog = saveFileDialog;
            _saveFileDialog.OverwritePrompt = true;

            _messageBox = messageBox;
            _projectsProvider = projectsProvider;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel))
            {
                return false;
            }

            var node = (CodeExplorerComponentViewModel)parameter;
            var componentType = node.Declaration.QualifiedName.QualifiedModuleName.ComponentType;

            if (componentType == ComponentType.Document)
            {
                return false;
            }

            return _exportableFileExtensions.Select(s => s.Key).Contains(componentType);
        }

        protected override void OnExecute(object parameter)
        {
            var message = string.Format(CodeExplorerUI.ExportBeforeRemove_Prompt, ((CodeExplorerComponentViewModel)parameter).Name);
            var result = _messageBox.Confirm(message, CodeExplorerUI.ExportBeforeRemove_Caption, ConfirmationOutcome.Yes);

            if (result == ConfirmationOutcome.Cancel)
            {
                return;
            }

            if (result == ConfirmationOutcome.Yes && !ExportFile((CodeExplorerComponentViewModel)parameter))
            {
                return;
            }

            // No file export or file successfully exported--now remove it

            // I know this will never be null because of the CanExecute
            var declaration = ((CodeExplorerComponentViewModel)parameter).Declaration;
            var qualifiedModuleName = declaration.QualifiedName.QualifiedModuleName;
            var components = _projectsProvider.ComponentsCollection(qualifiedModuleName.ProjectId);
            components?.Remove(_projectsProvider.Component(qualifiedModuleName));
        }

        private bool ExportFile(CodeExplorerComponentViewModel node)
        {
            var component = _projectsProvider.Component(node.Declaration.QualifiedName.QualifiedModuleName);

            _exportableFileExtensions.TryGetValue(component.Type, out string ext);

            _saveFileDialog.FileName = component.Name + ext;
            var result = _saveFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                component.Export(_saveFileDialog.FileName);
            }

            return result == DialogResult.OK;
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
