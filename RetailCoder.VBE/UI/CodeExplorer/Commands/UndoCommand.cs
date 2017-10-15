using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;
using Rubberduck.UI.SourceControl;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class UndoCommand : CommandBase
    {
        private readonly IDockablePresenter _presenter;
        private readonly IMessageBox _messageBox;

        public UndoCommand(IDockablePresenter presenter, IMessageBox messageBox) : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
            _messageBox = messageBox;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            var node = parameter as CodeExplorerComponentViewModel;
            if (node == null)
            {
                return false;
            }

            var panel = _presenter.UserControl as SourceControlPanel;
            Debug.Assert(panel != null);

            var panelViewModel = panel.ViewModel;
            if (panelViewModel == null)
            {
                return false;
            }

            panelViewModel.SetTab(SourceControlTab.Changes);
            var viewModel = panelViewModel.SelectedItem.ViewModel as ChangesViewViewModel;

            return viewModel != null && viewModel.IncludedChanges != null &&
                   viewModel.IncludedChanges.Select(s => s.FilePath).Contains(GetFileName(node));
        }

        protected override void ExecuteImpl(object parameter)
        {
            var panel = _presenter.UserControl as SourceControlPanel;
            Debug.Assert(panel != null);

            var panelViewModel = panel.ViewModel;
            if (panelViewModel == null)
            {
                return;
            }

            panelViewModel.SetTab(SourceControlTab.Changes);
            var viewModel = panelViewModel.SelectedItem.ViewModel as ChangesViewViewModel;
            if (viewModel == null)
            {
                return;
            }

            var fileName = GetFileName((ICodeExplorerDeclarationViewModel)parameter);
            var result = _messageBox.Show(string.Format(RubberduckUI.SourceControl_UndoPrompt, fileName),
                RubberduckUI.SourceControl_UndoTitle, System.Windows.Forms.MessageBoxButtons.OKCancel,
                System.Windows.Forms.MessageBoxIcon.Warning, System.Windows.Forms.MessageBoxDefaultButton.Button2);

            if (result != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            viewModel.UndoChangesToolbarButtonCommand.Execute(new FileStatusEntry(fileName, FileStatus.Modified));
            _presenter.Show();
        }

        private string GetFileName(ICodeExplorerDeclarationViewModel node)
        {
            var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;

            var fileExtensions = new Dictionary<ComponentType, string>
            {
                { ComponentType.StandardModule, ".bas" },
                { ComponentType.ClassModule, ".cls" },
                { ComponentType.Document, ".cls" },
                { ComponentType.UserForm, ".frm" }
            };

            string ext;
            fileExtensions.TryGetValue(component.Type, out ext);
            return component.Name + ext;
        }
    }
}