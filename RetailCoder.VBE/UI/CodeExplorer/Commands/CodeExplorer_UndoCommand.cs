using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;
using Rubberduck.UI.SourceControl;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_UndoCommand : CommandBase
    {
        private readonly SourceControlDockablePresenter _presenter;
        private readonly IMessageBox _messageBox;

        public CodeExplorer_UndoCommand(SourceControlDockablePresenter presenter, IMessageBox messageBox)
        {
            _presenter = presenter;
            _messageBox = messageBox;
        }

        public override bool CanExecute(object parameter)
        {
            var node = parameter as CodeExplorerComponentViewModel;
            if (node == null)
            {
                return false;
            }

            var panel = _presenter.Window() as SourceControlPanel;
            if (panel == null)
            {
                return false;
            }

            var panelVM = panel.ViewModel as SourceControlViewViewModel;
            if (panelVM == null)
            {
                return false;
            }

            panelVM.SetTab(SourceControlTab.Changes);
            var changesVM = panelVM.SelectedItem.ViewModel as ChangesViewViewModel;

            return changesVM != null && changesVM.IncludedChanges != null &&
                   changesVM.IncludedChanges.Select(s => s.FilePath).Contains(GetFileName(node));
        }

        public override void Execute(object parameter)
        {
            var panel = _presenter.Window() as SourceControlPanel;
            if (panel == null)
            {
                return;
            }

            var panelVM = panel.ViewModel as SourceControlViewViewModel;
            if (panelVM == null)
            {
                return;
            }

            panelVM.SetTab(SourceControlTab.Changes);
            var changesVM = panelVM.SelectedItem.ViewModel as ChangesViewViewModel;
            if (changesVM == null)
            {
                return;
            }

            var fileName = GetFileName((CodeExplorerComponentViewModel) parameter);
            var result = _messageBox.Show(string.Format(RubberduckUI.SourceControl_UndoPrompt, fileName),
                RubberduckUI.SourceControl_UndoTitle, System.Windows.Forms.MessageBoxButtons.OKCancel,
                System.Windows.Forms.MessageBoxIcon.Warning, System.Windows.Forms.MessageBoxDefaultButton.Button2);

            if (result != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            changesVM.UndoChangesToolbarButtonCommand.Execute(new FileStatusEntry(fileName, FileStatus.Modified));
            _presenter.Show();
        }

        private string GetFileName(CodeExplorerComponentViewModel node)
        {
            var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;

            var fileExtensions = new Dictionary<vbext_ComponentType, string>
            {
                { vbext_ComponentType.vbext_ct_StdModule, ".bas" },
                { vbext_ComponentType.vbext_ct_ClassModule, ".cls" },
                { vbext_ComponentType.vbext_ct_Document, ".cls" },
                { vbext_ComponentType.vbext_ct_MSForm, ".frm" }
            };

            string ext;
            fileExtensions.TryGetValue(component.Type, out ext);
            return component.Name + ext;
        }
    }
}