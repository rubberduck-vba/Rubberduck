using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_RefreshComponentCommand : CommandBase
    {
        private readonly RubberduckParserState _state;

        public CodeExplorer_RefreshComponentCommand(RubberduckParserState state)
        {
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            var node = parameter as CodeExplorerComponentViewModel;

            return node != null && node.QualifiedSelection.HasValue &&
                   _state.GetOrCreateModuleState(node.QualifiedSelection.Value.QualifiedName.Component) == ParserState.Pending;
        }

        public override void Execute(object parameter)
        {
            var node = (CodeExplorerComponentViewModel) parameter;

            // ReSharper disable once PossibleInvalidOperationException - CanExecute ensures it has a value
            _state.OnParseRequested(this, node.QualifiedSelection.Value.QualifiedName.Component);
        }
    }
}