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

        public CodeExplorer_UndoCommand(SourceControlDockablePresenter presenter)
        {
            _presenter = presenter;
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

            return changesVM != null && changesVM.IncludedChanges.Select(s => s.FilePath).Contains(GetFileName(node));
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();

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

            changesVM.UndoChangesToolbarButtonCommand.Execute(new FileStatusEntry(GetFileName((CodeExplorerComponentViewModel)parameter), FileStatus.Modified));
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