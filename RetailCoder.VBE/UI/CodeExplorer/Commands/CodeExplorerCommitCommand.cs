using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.UI.SourceControl;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerCommitCommand : CommandBase
    {
        private readonly SourceControlDockablePresenter _presenter;

        public CodeExplorerCommitCommand(SourceControlDockablePresenter presenter) : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return parameter is CodeExplorerComponentViewModel;
        }

        protected override void ExecuteImpl(object parameter)
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
