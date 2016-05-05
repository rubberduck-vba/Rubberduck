using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerFindAllReferencesCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly FindAllReferencesCommand _findAllReferences;

        public CodeExplorerFindAllReferencesCommand(RubberduckParserState state, FindAllReferencesCommand findAllReferences)
        {
            _state = state;
            _findAllReferences = findAllReferences;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready && !(parameter is CodeExplorerCustomFolderViewModel) &&
                !(parameter is CodeExplorerErrorNodeViewModel);
        }

        public override void Execute(object parameter)
        {
            _findAllReferences.Execute(GetSelectedDeclaration((CodeExplorerItemViewModel) parameter));
        }

        private Declaration GetSelectedDeclaration(CodeExplorerItemViewModel node)
        {
            if (node is CodeExplorerProjectViewModel)
            {
                return ((CodeExplorerProjectViewModel)node).Declaration;
            }

            if (node is CodeExplorerComponentViewModel)
            {
                return ((CodeExplorerComponentViewModel)node).Declaration;
            }

            if (node is CodeExplorerMemberViewModel)
            {
                return ((CodeExplorerMemberViewModel)node).Declaration;
            }

            return null;
        }
    }
}