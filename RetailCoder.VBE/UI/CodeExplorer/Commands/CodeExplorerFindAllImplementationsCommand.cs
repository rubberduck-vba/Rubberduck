using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerRindAllImplementationsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly FindAllImplementationsCommand _findAllImplementations;

        public CodeExplorerRindAllImplementationsCommand(RubberduckParserState state, FindAllImplementationsCommand findAllImplementations)
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
            _findAllImplementations.Execute(GetSelectedDeclaration((CodeExplorerItemViewModel) parameter));
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