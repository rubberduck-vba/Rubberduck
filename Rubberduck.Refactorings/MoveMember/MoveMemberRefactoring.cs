using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;
using System;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringTestAccess
    {
        MoveMemberModel TestUserInteractionOnly(Declaration target, Func<MoveMemberModel, MoveMemberModel> userInteraction);
        string PreviewModuleContent(MoveMemberModel model, PreviewModule previewModule);
    }

    public class MoveMemberRefactoring : InteractiveRefactoringBase<IMoveMemberPresenter, MoveMemberModel>, IMoveMemberRefactoringTestAccess
    {
        private readonly MoveMemberRefactoringAction _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRewritingManager _rewritingManager;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        private MoveMemberObjectsFactory _moveMemberFactory;

        public MoveMemberRefactoring(
            MoveMemberRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            IRefactoringPresenterFactory factory,
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IUiDispatcher uiDispatcher)
                : base(selectionProvider, factory, uiDispatcher)

        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _moveMemberFactory = new MoveMemberObjectsFactory(declarationFinderProvider);
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selected = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selected.IsMember()
                || selected.IsModuleConstant()
                || (selected.IsField() && !selected.HasPrivateAccessibility()))
            {
                return selected;
            }

            return null;
        }

        public MoveMemberModel TestUserInteractionOnly(Declaration target, Func<MoveMemberModel, MoveMemberModel> userInteraction)
        {
            var model = InitializeModel(target);
            return userInteraction(model);
        }

        public string PreviewModuleContent(MoveMemberModel model, PreviewModule previewModule)
        {
            if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy))
            {
                return MoveMemberResources.ApplicableStrategyNotFound;
            }

            var isExistingDestination = model.Destination.IsExistingModule(out var destinationModule);
            if (previewModule == PreviewModule.Destination && !isExistingDestination)
            {
                var content = strategy.NewDestinationModuleContent(model, _rewritingManager, new MovedContentProvider()).AsSingleBlockWithinDemarcationComments();

                return $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}{Environment.NewLine}{content}";
            }

            var previewSession = _rewritingManager.CheckOutCodePaneSession();
            strategy.RefactorRewrite(model, previewSession, _rewritingManager, true);

            var qmnToPreview = previewModule == PreviewModule.Destination
                ? destinationModule.QualifiedModuleName
                : model.Source.QualifiedModuleName;

            var rewriter = previewSession.CheckOutModuleRewriter(qmnToPreview);
            var preview = rewriter.GetText(maxConsecutiveNewLines: 3);
            return preview;
        }

        protected override MoveMemberModel InitializeModel(Declaration target)
        {
            if (target == null) { throw new TargetDeclarationIsNullException(); }

            var model = new MoveMemberModel(target, _declarationFinderProvider, PreviewModuleContent, _moveMemberFactory);
            return model;
        }

        protected override void RefactorImpl(MoveMemberModel model)
        {
            _refactoringAction.Refactor(model);
        }
    }
}
