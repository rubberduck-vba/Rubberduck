using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class RemoveCommand : CommandBase, IDisposable
    {
        private readonly ISaveFileDialog _saveFileDialog;
        private readonly IMessageBox _messageBox;

        private readonly Dictionary<ComponentType, string> _exportableFileExtensions = new Dictionary<ComponentType, string>
        {
            { ComponentType.StandardModule, ".bas" },
            { ComponentType.ClassModule, ".cls" },
            { ComponentType.Document, ".cls" },
            { ComponentType.UserForm, ".frm" }
        };

        public RemoveCommand(ISaveFileDialog saveFileDialog, IMessageBox messageBox) : base(LogManager.GetCurrentClassLogger())
        {
            _saveFileDialog = saveFileDialog;
            _saveFileDialog.OverwritePrompt = true;

            _messageBox = messageBox;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel))
            {
                return false;
            }

            var node = (CodeExplorerComponentViewModel)parameter;
            var componentType = node.Declaration.QualifiedName.QualifiedModuleName.ComponentType;
            return _exportableFileExtensions.Select(s => s.Key).Contains(componentType);
        }

        protected override void ExecuteImpl(object parameter)
        {
            var message = string.Format("Do you want to export '{0}' before removing?", ((CodeExplorerComponentViewModel)parameter).Name);
            var result = _messageBox.Show(message, "Rubberduck Export Prompt", MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);

            if (result == DialogResult.Cancel)
            {
                return;
            }

            if (result == DialogResult.Yes && !ExportFile((CodeExplorerComponentViewModel)parameter))
            {
                return;
            }

            // No file export or file successfully exported--now remove it

            // I know this will never be null because of the CanExecute
            var declaration = ((CodeExplorerComponentViewModel)parameter).Declaration;

            var components = declaration.Project.VBComponents;
            components.Remove(declaration.QualifiedName.QualifiedModuleName.Component);
        }

        private bool ExportFile(CodeExplorerComponentViewModel node)
        {
            var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;

            string ext;
            _exportableFileExtensions.TryGetValue(component.Type, out ext);

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
