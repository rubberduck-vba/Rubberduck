using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class IndentCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private readonly INavigateCommand _navigateCommand;

        public IndentCommand(RubberduckParserState state, IIndenter indenter, INavigateCommand navigateCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _indenter = indenter;
            _navigateCommand = navigateCommand;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            if (parameter == null)
            {
                return false;
            }

            var model = parameter as CodeExplorerComponentViewModel;
            if (model != null)
            {
                var node = model;
                if (node.Declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.NoIndent))
                {
                    return false;
                }
            }

            if (parameter is CodeExplorerProjectViewModel)
            {
                if (_state.Status != ParserState.Ready)
                {
                    return false;
                }

                var declaration = ((ICodeExplorerDeclarationViewModel)parameter).Declaration;
                return _state.AllUserDeclarations
                            .Any(c => c.DeclarationType.HasFlag(DeclarationType.Module) &&
                            c.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent) &&
                            c.Project == declaration.Project);
            }

            if (parameter is CodeExplorerCustomFolderViewModel)
            {
                if (_state.Status != ParserState.Ready)
                {
                    return false;
                }

                var node = (CodeExplorerCustomFolderViewModel) parameter;
                return node.Items.OfType<CodeExplorerComponentViewModel>()
                        .Select(s => s.Declaration)
                        .Any(d => d.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent));
            }

            return _state.Status == ParserState.Ready;
        }

        protected override void ExecuteImpl(object parameter)
        {
            if (parameter == null)
            {
                return;
            }

            var node = (CodeExplorerItemViewModel)parameter;

            if (!node.QualifiedSelection.HasValue && !(node is CodeExplorerCustomFolderViewModel))
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

            if (node is CodeExplorerCustomFolderViewModel)
            {
                var components = node.Items.OfType<CodeExplorerComponentViewModel>()
                        .Select(s => s.Declaration)
                        .Where(d => d.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent))
                        .Select(d => d.QualifiedName.QualifiedModuleName.Component);

                foreach (var component in components)
                {
                    _indenter.Indent(component);
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