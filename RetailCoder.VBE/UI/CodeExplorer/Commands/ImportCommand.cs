using System;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class ImportCommand : CommandBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly IOpenFileDialog _openFileDialog;

        public ImportCommand(IVBE vbe, IOpenFileDialog openFileDialog) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _openFileDialog = openFileDialog;

            _openFileDialog.AddExtension = true;
            _openFileDialog.AutoUpgradeEnabled = true;
            _openFileDialog.CheckFileExists = true;
            _openFileDialog.Multiselect = true;
            _openFileDialog.ShowHelp = false;   // we don't want 1996's file picker.
            _openFileDialog.Filter = @"VB Files|*.cls;*.bas;*.frm";
            _openFileDialog.CheckFileExists = true;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return parameter != null || _vbe.VBProjects.Count == 1 || _vbe.ActiveVBProject != null;
        }

        protected override void ExecuteImpl(object parameter)
        {
            var project = GetNodeProject((CodeExplorerItemViewModel)parameter);
            if (project == null)
            {
                if (_vbe.VBProjects.Count == 1)
                {
                    project = _vbe.VBProjects[1];
                }
                else if (_vbe.ActiveVBProject != null)
                {
                    project = _vbe.ActiveVBProject;
                }
            }

            if (project == null || _openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

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

        private IVBProject GetNodeProject(CodeExplorerItemViewModel parameter)
        {
            if (parameter is ICodeExplorerDeclarationViewModel)
            {
                return parameter.GetSelectedDeclaration().Project;
            }

            var node = parameter.Parent;
            while (!(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return ((ICodeExplorerDeclarationViewModel)node).Declaration.Project;
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
