using System.Linq;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_ImportCommand : CommandBase
    {
        private readonly OpenFileDialog _openFileDialog;

        public CodeExplorer_ImportCommand(OpenFileDialog openFileDialog)
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
            return parameter is CodeExplorerProjectViewModel ||
                   parameter is CodeExplorerComponentViewModel ||
                   parameter is CodeExplorerMemberViewModel;
        }

        public override void Execute(object parameter)
        {
            // I know this will never be null because of the CanExecute
            var project = GetSelectedDeclaration((CodeExplorerItemViewModel)parameter).QualifiedName.QualifiedModuleName.Project;

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

        private Declaration GetSelectedDeclaration(CodeExplorerItemViewModel node)
        {
            if (node is CodeExplorerProjectViewModel)
            {
                return ((CodeExplorerProjectViewModel)node).Declaration;
            }

            if (node is CodeExplorerComponentViewModel)
            {
                return ((CodeExplorerComponentViewModel)node).Declaration;
            }

            if (node is CodeExplorerMemberViewModel)
            {
                return ((CodeExplorerMemberViewModel)node).Declaration;
            }

            return null;
        }
    }
}