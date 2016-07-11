using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerNavigateCommand : CommandBase
    {
        private readonly INavigateCommand _navigateCommand;

        public CodeExplorerNavigateCommand(INavigateCommand navigateCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _navigateCommand = navigateCommand;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return parameter != null && ((CodeExplorerItemViewModel)parameter).QualifiedSelection.HasValue;
        }

        protected override void ExecuteImpl(object parameter)
        {
            // ReSharper disable once PossibleInvalidOperationException
            _navigateCommand.Execute(((CodeExplorerItemViewModel)parameter).QualifiedSelection.Value.GetNavitationArgs());
        }
    }
}
