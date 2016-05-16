using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_FindAllImplementationsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly FindAllImplementationsCommand _findAllImplementations;

        public CodeExplorer_FindAllImplementationsCommand(RubberduckParserState state, FindAllImplementationsCommand findAllImplementations)
        {
            _state = state;
            _findAllImplementations = findAllImplementations;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready &&
                   (parameter is CodeExplorerComponentViewModel ||
                    parameter is CodeExplorerMemberViewModel);
        }

        public override void Execute(object parameter)
        {
            _findAllImplementations.Execute(((CodeExplorerItemViewModel) parameter).GetSelectedDeclaration());
        }
    }
}