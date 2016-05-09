using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.UI.SourceControl;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_CommitCommand : CommandBase
    {
        private readonly SourceControlPanel _panel;

        public CodeExplorer_CommitCommand(SourceControlPanel panel)
        {
            _panel = panel;
        }

        public override bool CanExecute(object parameter)
        {
            return parameter is CodeExplorerComponentViewModel;
        }

        public override void Execute(object parameter)
        {
            _panel.Show();

            var vm = _panel.ViewModel as SourceControlViewViewModel;

            if (vm != null)
            {
                vm.SetTab(SourceControlTab.Changes);
            }
        }
    }
}