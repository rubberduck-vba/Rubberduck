using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_RefreshComponentCommand : CommandBase
    {
        private readonly RubberduckParserState _state;

        public CodeExplorer_RefreshComponentCommand(RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            var node = parameter as CodeExplorerComponentViewModel;

            return node != null && node.QualifiedSelection.HasValue &&
                   _state.GetOrCreateModuleState(node.QualifiedSelection.Value.QualifiedName.Component) == ParserState.Pending;
        }

        protected override void ExecuteImpl(object parameter)
        {
            var node = (CodeExplorerComponentViewModel) parameter;

            // ReSharper disable once PossibleInvalidOperationException - CanExecute ensures it has a value
            _state.OnParseRequested(this, node.QualifiedSelection.Value.QualifiedName.Component);
        }
    }
}
