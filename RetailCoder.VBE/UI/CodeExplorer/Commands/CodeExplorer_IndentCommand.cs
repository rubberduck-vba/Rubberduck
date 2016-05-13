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

            if (parameter is CodeExplorerProjectViewModel)
            {
                return _state.Status == ParserState.Ready &&
                    _state.AllUserDeclarations.Any(c =>
                            c.DeclarationType.HasFlag(DeclarationType.Module) &&
                            c.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent));
            }

            return _state.Status == ParserState.Ready && !(parameter is CodeExplorerErrorNodeViewModel);
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
                var declaration = ((ICodeExplorerDeclarationViewModel)node).Declaration;

                var components = _state.AllUserDeclarations.Where(c => 
                            c.DeclarationType.HasFlag(DeclarationType.Module) &&
                            c.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent) &&
                            c.Project == declaration.Project);

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