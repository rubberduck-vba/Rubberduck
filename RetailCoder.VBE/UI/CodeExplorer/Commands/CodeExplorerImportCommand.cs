using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerImportCommand : CommandBase, IDisposable
    {
        private readonly IOpenFileDialog _openFileDialog;

        public CodeExplorerImportCommand(IOpenFileDialog openFileDialog) : base(LogManager.GetCurrentClassLogger())
        {
            _openFileDialog = openFileDialog;

            _openFileDialog.AddExtension = true;
            _openFileDialog.AutoUpgradeEnabled = true;
            _openFileDialog.CheckFileExists = true;
            _openFileDialog.Multiselect = true;
            _openFileDialog.ShowHelp = false;   // we don't want 1996's file picker.
            _openFileDialog.Filter = @"VB Files|*.cls;*.bas;*.frm";
            _openFileDialog.CheckFileExists = true;
        }

        protected override void ExecuteImpl(object parameter)
        {
            VBProject project;

            if (parameter is ICodeExplorerDeclarationViewModel)
            {
                project = ((ICodeExplorerDeclarationViewModel) parameter).Declaration.QualifiedName.QualifiedModuleName.Project;
            }
            else
            {
                var node = ((CodeExplorerItemViewModel) parameter).Parent;
                while (!(node is ICodeExplorerDeclarationViewModel))
                {
                    node = node.Parent;
                }

                project = ((ICodeExplorerDeclarationViewModel) node).Declaration.QualifiedName.QualifiedModuleName.Project;
            }

            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var fileExts = _openFileDialog.FileNames.Select(s => s.Split('.').Last());
                if (fileExts.Any(fileExt => !new[] {"bas", "cls", "frm"}.Contains(fileExt)))
                {
                    return;
                }

                foreach (var filename in _openFileDialog.FileNames)
                {
                    project.VBComponents.Import(filename);
                }
            }
        }

        public void Dispose()
        {
            if (_openFileDialog != null)
            {
                _openFileDialog.Dispose();
            }
        }
    }
}
