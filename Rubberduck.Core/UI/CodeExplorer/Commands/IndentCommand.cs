using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Interaction.Navigation;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class IndentCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private readonly INavigateCommand _navigateCommand;

        public IndentCommand(RubberduckParserState state, IIndenter indenter, INavigateCommand navigateCommand)
        {
            _state = state;
            _indenter = indenter;
            _navigateCommand = navigateCommand;
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!base.EvaluateCanExecute(parameter) || _state.Status != ParserState.Ready)
            {
                return false;
            }

            switch (parameter)
            {
                case CodeExplorerProjectViewModel project:
                    return _state.AllUserDeclarations
                        .Any(c => c.DeclarationType.HasFlag(DeclarationType.Module) &&
                                  c.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent) &&
                                  c.ProjectId == project.Declaration.ProjectId);
                case CodeExplorerCustomFolderViewModel folder:
                    return folder.Children.OfType<CodeExplorerComponentViewModel>()     //TODO - this has the filter applied.
                        .Select(s => s.Declaration)
                        .Any(d => d.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent));
                case CodeExplorerComponentViewModel model:
                    return model.Declaration.Annotations.Any(a => a.AnnotationType != AnnotationType.NoIndent);
                case CodeExplorerMemberViewModel member:
                    return member.QualifiedSelection.HasValue; 
                default:
                    return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            if (!base.EvaluateCanExecute(parameter) || 
                !(parameter is CodeExplorerItemViewModel node) ||
                node.Declaration == null)
            {
                return;
            }

            switch (node)
            {
                case CodeExplorerProjectViewModel model:
                {
                    var declaration = model.Declaration;

                    var componentDeclarations = _state.AllUserDeclarations.Where(c => 
                        c.DeclarationType.HasFlag(DeclarationType.Module) &&
                        c.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent) &&
                        c.ProjectId == declaration.ProjectId);

                    foreach (var componentDeclaration in componentDeclarations)
                    {
                        _indenter.Indent(_state.ProjectsProvider.Component(componentDeclaration.QualifiedName.QualifiedModuleName));
                    }

                    break;
                }
                case CodeExplorerCustomFolderViewModel folder:
                {
                    var components = folder.Children.OfType<CodeExplorerComponentViewModel>()   //TODO: this has the filter applied.
                        .Select(s => s.Declaration)
                        .Where(d => d.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent))
                        .Select(d => _state.ProjectsProvider.Component(d.QualifiedName.QualifiedModuleName));

                    foreach (var component in components)
                    {
                        _indenter.Indent(component);
                    }

                    break;
                }
                case CodeExplorerComponentViewModel component:
                    _indenter.Indent(_state.ProjectsProvider.Component(component.Declaration.QualifiedModuleName));
                    break;
                case CodeExplorerMemberViewModel member:
                    if (!member.QualifiedSelection.HasValue)
                    {
                        return;
                    }
                    _navigateCommand.Execute(member.QualifiedSelection.Value.GetNavitationArgs());
                    _indenter.IndentCurrentProcedure();
                    break;
            }
        }
    }
}