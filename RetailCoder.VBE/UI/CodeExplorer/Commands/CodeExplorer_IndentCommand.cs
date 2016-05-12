using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;

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
            if (parameter is CodeExplorerComponentViewModel)
            {
                var node = (CodeExplorerComponentViewModel)parameter;
                if (node.Declaration.Annotations.Any(a => a.AnnotationType == Parsing.Annotations.AnnotationType.NoIndent))
                {
                    return false;
                }
            }

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
                var moduleTypes = new[] {DeclarationType.ClassModule, DeclarationType.Module};
                var components = _state.AllUserDeclarations.Where(c =>
                            moduleTypes.Contains(c.DeclarationType) &&
                            c.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent));

                foreach (var component in components)
                {
                    _indenter.Indent(component.QualifiedName.QualifiedModuleName.Component);
                }
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