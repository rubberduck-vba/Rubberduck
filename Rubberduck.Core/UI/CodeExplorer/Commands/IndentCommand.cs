using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using System.Linq;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Interaction.Navigation;
using Rubberduck.VBEditor.Events;

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

        public IndentCommand(
            RubberduckParserState state, 
            IIndenter indenter, 
            INavigateCommand navigateCommand, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _state = state;
            _indenter = indenter;
            _navigateCommand = navigateCommand;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            switch (parameter)
            {
                case CodeExplorerProjectViewModel project:
                    return _state.AllUserDeclarations
                        .Any(c => c.DeclarationType.HasFlag(DeclarationType.Module) &&
                                  !c.Annotations.Any(pta => pta.Annotation is NoIndentAnnotation) &&
                                  c.ProjectId == project.Declaration.ProjectId);
                case CodeExplorerCustomFolderViewModel folder:
                    return folder.Children.OfType<CodeExplorerComponentViewModel>()     //TODO - this has the filter applied.
                        .Select(s => s.Declaration)
                        .Any(d => !d.Annotations.Any(pta => pta.Annotation is NoIndentAnnotation));
                case CodeExplorerComponentViewModel model:
                    return !model.Declaration.Annotations.Any(pta => pta.Annotation is NoIndentAnnotation);
                case CodeExplorerMemberViewModel member:
                    return member.QualifiedSelection.HasValue; 
                default:
                    return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter) || 
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
                        !c.Annotations.Any(pta => pta.Annotation is NoIndentAnnotation) &&
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
                        .Where(d => !d.Annotations.Any(pta => pta.Annotation is NoIndentAnnotation))
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