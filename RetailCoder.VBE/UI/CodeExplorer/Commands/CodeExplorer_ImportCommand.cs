using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_ImportCommand : CommandBase, IDisposable
    {
        private readonly IOpenFileDialog _openFileDialog;

        public CodeExplorer_ImportCommand(IOpenFileDialog openFileDialog)
        {
            _openFileDialog = openFileDialog;

            _openFileDialog.AddExtension = true;
            _openFileDialog.AutoUpgradeEnabled = true;
            _openFileDialog.CheckFileExists = true;
            _openFileDialog.Multiselect = false;
            _openFileDialog.ShowHelp = false;   // we don't want 1996's file picker.
            _openFileDialog.Filter = @"VB Files|*.cls;*.bas;*.frm";
            _openFileDialog.CheckFileExists = true;
        }

        public override bool CanExecute(object parameter)
        {
            // I could import to a folder as well, if I had a
            // MoveToFolder refactoring to call
            return parameter is ICodeExplorerDeclarationViewModel;
        }

        public override void Execute(object parameter)
        {
            // I know this will never be null because of the CanExecute
            var project = ((ICodeExplorerDeclarationViewModel)parameter).Declaration.QualifiedName.QualifiedModuleName.Project;

            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var fileExt = _openFileDialog.FileName.Split('.').Last();
                if (!new[]{"bas", "cls", "frm"}.Contains(fileExt))
                {
                    return;
                }

                project.VBComponents.Import(_openFileDialog.FileName);
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