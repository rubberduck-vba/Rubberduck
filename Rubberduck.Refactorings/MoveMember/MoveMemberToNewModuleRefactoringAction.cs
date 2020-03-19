using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.Utility;
using System;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberToNewModuleRefactoringAction : RefactoringActionWithSuspension<MoveMemberModel>
    {
        private readonly MoveMemberExistingModulesRefactoringAction _refactoring;
        private readonly IRewritingManager _rewritingManager;
        private readonly IAddComponentService _addComponentService;

        public MoveMemberToNewModuleRefactoringAction(
                        MoveMemberExistingModulesRefactoringAction refactoring,
                        RenameCodeDefinedIdentifierRefactoringAction renameAction,
                        IParseManager parseManager,
                        IRewritingManager rewritingManager,
                        IAddComponentService addComponentService)
                : base(parseManager, rewritingManager)
        {
            _refactoring = refactoring;
            _rewritingManager = rewritingManager;
            _addComponentService = addComponentService;
        }

        protected override void Refactor(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy) 
                || !strategy.IsExecutableModel(model, out _))
            {
                return;
            }

            var optionExplicit = $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}";

            var newContent = strategy.NewDestinationModuleContent(model, _rewritingManager, new MovedContentProvider()).AsSingleBlock;

            _refactoring.Refactor(model, rewriteSession);

            _addComponentService.AddComponentWithAttributes(
                                        model.Source.Module.ProjectId,
                                        model.Destination.ComponentType,
                                        $"{optionExplicit}{Environment.NewLine}{newContent}",
                                        componentName: model.Destination.ModuleName);
        }

        protected override bool RequiresSuspension(MoveMemberModel model) => true;
    }
}
