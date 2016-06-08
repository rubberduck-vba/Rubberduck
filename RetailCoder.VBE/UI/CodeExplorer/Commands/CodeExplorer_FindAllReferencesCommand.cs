using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_FindAllReferencesCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly FindAllReferencesCommand _findAllReferences;

        public CodeExplorer_FindAllReferencesCommand(RubberduckParserState state, FindAllReferencesCommand findAllReferences)
        {
            _state = state;
            _findAllReferences = findAllReferences;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready &&
                parameter != null &&
                !(parameter is CodeExplorerCustomFolderViewModel);
        }

        public override void Execute(object parameter)
        {
            _findAllReferences.Execute(((CodeExplorerItemViewModel) parameter).GetSelectedDeclaration());
        }
    }
}
