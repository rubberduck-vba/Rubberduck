using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_NavigateCommand : CommandBase
    {
        private readonly INavigateCommand _navigateCommand;

        public CodeExplorer_NavigateCommand(INavigateCommand navigateCommand)
        {
            _navigateCommand = navigateCommand;
        }

        public override bool CanExecute(object parameter)
        {
            return parameter != null && ((CodeExplorerItemViewModel)parameter).QualifiedSelection.HasValue;
        }

        public override void Execute(object parameter)
        {
            // ReSharper disable once PossibleInvalidOperationException
            _navigateCommand.Execute(((CodeExplorerItemViewModel)parameter).QualifiedSelection.Value.GetNavitationArgs());
        }
    }
}
