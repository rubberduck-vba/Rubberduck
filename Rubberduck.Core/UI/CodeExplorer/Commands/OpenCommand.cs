using NLog;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class OpenCommand : CommandBase
    {
        private readonly INavigateCommand _openCommand;

        public OpenCommand(INavigateCommand OpenCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _openCommand = OpenCommand;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return parameter != null && ((CodeExplorerItemViewModel)parameter).QualifiedSelection.HasValue;
        }

        protected override void OnExecute(object parameter)
        {
            // ReSharper disable once PossibleInvalidOperationException
            _openCommand.Execute(((CodeExplorerItemViewModel)parameter).QualifiedSelection.Value.GetNavitationArgs());
        }
    }
}
