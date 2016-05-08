using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_IndentCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private readonly INavigateCommand _navigateCommand;

        public CodeExplorer_IndentCommand(RubberduckParserState state, IIndenter indenter, INavigateCommand navigateCommand)
        {
            _state = state;
            _indenter = indenter;
            _navigateCommand = navigateCommand;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready && !(parameter is CodeExplorerCustomFolderViewModel) &&
                   !(parameter is CodeExplorerErrorNodeViewModel);
        }

        public override void Execute(object parameter)
        {
            var node = (CodeExplorerItemViewModel)parameter;

            if (!node.QualifiedSelection.HasValue)
            {
                return;
            }

            if (node is CodeExplorerProjectViewModel)
            {
                _indenter.Indent(node.QualifiedSelection.Value.QualifiedName.Project);
            }

            if (node is CodeExplorerComponentViewModel)
            {
                _indenter.Indent(node.QualifiedSelection.Value.QualifiedName.Component);
            }

            if (node is CodeExplorerMemberViewModel)
            {
                _navigateCommand.Execute(node.QualifiedSelection.Value.GetNavitationArgs());

                _indenter.IndentCurrentProcedure();
            }
        }
    }
}