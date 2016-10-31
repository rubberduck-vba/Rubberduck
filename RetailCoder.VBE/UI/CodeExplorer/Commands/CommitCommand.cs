using System.Diagnostics;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.UI.SourceControl;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class CommitCommand : CommandBase
    {
        private readonly IDockablePresenter _presenter;

        public CommitCommand(IDockablePresenter presenter) : base(LogManager.GetCurrentClassLogger())
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

            var panel = _presenter.UserControl as SourceControlPanel;
            Debug.Assert(panel != null);

            var vm = panel.ViewModel;
            if (vm != null)
            {
                vm.SetTab(SourceControlTab.Changes);
            }
        }
    }
}
