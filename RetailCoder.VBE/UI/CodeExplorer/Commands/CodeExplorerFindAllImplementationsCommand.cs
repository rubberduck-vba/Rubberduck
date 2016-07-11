using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerFindAllImplementationsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly FindAllImplementationsCommand _findAllImplementations;

        public CodeExplorerFindAllImplementationsCommand(RubberduckParserState state, FindAllImplementationsCommand findAllImplementations) : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _findAllImplementations = findAllImplementations;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _state.Status == ParserState.Ready &&
                   (parameter is CodeExplorerComponentViewModel ||
                    parameter is CodeExplorerMemberViewModel);
        }

        protected override void ExecuteImpl(object parameter)
        {
            _findAllImplementations.Execute(((CodeExplorerItemViewModel) parameter).GetSelectedDeclaration());
        }
    }
}
