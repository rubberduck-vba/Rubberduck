using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_RemoveCommand : CommandBase, IDisposable
    {
        private readonly ISaveFileDialog _saveFileDialog;
        private readonly IMessageBox _messageBox;

        private readonly Dictionary<vbext_ComponentType, string> _exportableFileExtensions = new Dictionary<vbext_ComponentType, string>
        {
            { vbext_ComponentType.vbext_ct_StdModule, ".bas" },
            { vbext_ComponentType.vbext_ct_ClassModule, ".cls" },
            { vbext_ComponentType.vbext_ct_Document, ".cls" },
            { vbext_ComponentType.vbext_ct_MSForm, ".frm" }
        };

        public CodeExplorer_RemoveCommand(ISaveFileDialog saveFileDialog, IMessageBox messageBox)
        {
            _saveFileDialog = saveFileDialog;
            _saveFileDialog.OverwritePrompt = true;

            _messageBox = messageBox;
        }

        public override bool CanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel))
            {
                return false;
            }

            var node = (CodeExplorerComponentViewModel)parameter;
            var componentType = node.Declaration.QualifiedName.QualifiedModuleName.Component.Type;
            return _exportableFileExtensions.Select(s => s.Key).Contains(componentType);
        }

        public override void Execute(object parameter)
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

            var project = declaration.QualifiedName.QualifiedModuleName.Project;
            project.VBComponents.Remove(declaration.QualifiedName.QualifiedModuleName.Component);
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