using System;
using System.Collections.Generic;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public sealed class RenameCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly RubberduckParserState _state;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly IMessageBox _msgBox;
        private readonly IRewritingManager _rewritingManager;
        private readonly ISelectionService _selectionService;

        public RenameCommand(RubberduckParserState state, IMessageBox msgBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            _selectionService = selectionService;
            _state = state;
            _rewritingManager = rewritingManager;
            _msgBox = msgBox;
            _factory = factory;
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready && base.EvaluateCanExecute(parameter);
        }

        protected override void OnExecute(object parameter)
        {
            if (!EvaluateCanExecute(parameter) ||
                !(parameter is CodeExplorerItemViewModel node) ||
                node.Declaration == null)
            {
                return;
            }

            var refactoring = new RenameRefactoring(_factory, _msgBox, _state, _state.ProjectsProvider, _rewritingManager, _selectionService);
            refactoring.Refactor(node.Declaration);
        }
    }
}
