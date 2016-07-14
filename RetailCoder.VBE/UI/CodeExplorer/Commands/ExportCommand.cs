using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class ExportCommand : CommandBase, IDisposable
    {
        private readonly ISaveFileDialog _saveFileDialog;
        private readonly Dictionary<vbext_ComponentType, string> _exportableFileExtensions = new Dictionary<vbext_ComponentType, string>
        {
            { vbext_ComponentType.vbext_ct_StdModule, ".bas" },
            { vbext_ComponentType.vbext_ct_ClassModule, ".cls" },
            { vbext_ComponentType.vbext_ct_Document, ".cls" },
            { vbext_ComponentType.vbext_ct_MSForm, ".frm" }
        };

        public ExportCommand(ISaveFileDialog saveFileDialog) : base(LogManager.GetCurrentClassLogger())
        {
            _saveFileDialog = saveFileDialog;
            _saveFileDialog.OverwritePrompt = true;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel))
            {
                return false;
            }

            try
            {
                var node = (CodeExplorerComponentViewModel)parameter;
                var componentType = node.Declaration.QualifiedName.QualifiedModuleName.Component.Type;
                return _exportableFileExtensions.Select(s => s.Key).Contains(componentType);
            }
            catch (COMException)
            {
                // thrown when the component reference is stale
                return false;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            var node = (CodeExplorerComponentViewModel)parameter;
            var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;

            string ext;
            _exportableFileExtensions.TryGetValue(component.Type, out ext);

            _saveFileDialog.FileName = component.Name + ext;
            var result = _saveFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
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
