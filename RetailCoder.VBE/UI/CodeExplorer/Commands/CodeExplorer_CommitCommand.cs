using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.UI.SourceControl;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_CommitCommand : CommandBase
    {
        private readonly SourceControlDockablePresenter _presenter;

        public CodeExplorer_CommitCommand(SourceControlDockablePresenter presenter)
        {
            _presenter = presenter;
        }

        public override bool CanExecuteImpl(object parameter)
        {
            return parameter is CodeExplorerComponentViewModel;
        }

        public override void ExecuteImpl(object parameter)
        {
            _presenter.Show();

            var panel = _presenter.Window() as SourceControlPanel;
            if (panel != null)
            {
                var vm = panel.ViewModel as SourceControlViewViewModel;
                if (vm != null)
                {
                    vm.SetTab(SourceControlTab.Changes);
                }
            }
        }
    }
}
