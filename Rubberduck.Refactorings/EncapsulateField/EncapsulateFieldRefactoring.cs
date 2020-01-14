using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Utility;
using System.Collections.Generic;
using System;
using Rubberduck.Refactorings.EncapsulateField.Extensions;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public enum EncapsulateFieldStrategy
    {
        UseBackingFields,
        ConvertFieldsToUDTMembers
    }

    public interface IEncapsulateFieldRefactoringTestAccess
    {
        EncapsulateFieldModel TestUserInteractionOnly(Declaration target, Func<EncapsulateFieldModel, EncapsulateFieldModel> userInteraction);
    }

    public class EncapsulateFieldRefactoring : InteractiveRefactoringBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>, IEncapsulateFieldRefactoringTestAccess
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IIndenter _indenter;
        private QualifiedModuleName _targetQMN;

        public EncapsulateFieldRefactoring(
                IDeclarationFinderProvider declarationFinderProvider,
                IIndenter indenter,
                IRefactoringPresenterFactory factory,
                IRewritingManager rewritingManager,
                ISelectionProvider selectionProvider,
                ISelectedDeclarationProvider selectedDeclarationProvider,
                IUiDispatcher uiDispatcher)
            :base(rewritingManager, selectionProvider, factory, uiDispatcher)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _indenter = indenter;
        }

        public EncapsulateFieldModel TestUserInteractionOnly(Declaration target, Func<EncapsulateFieldModel, EncapsulateFieldModel> userInteraction)
        {
            var model = InitializeModel(target);
            return userInteraction(model);
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selectedDeclaration == null
                || selectedDeclaration.DeclarationType != DeclarationType.Variable
                || selectedDeclaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return null;
            }

            return selectedDeclaration;
        }

        protected override EncapsulateFieldModel InitializeModel(Declaration target)
        {
            if (target == null) { throw new TargetDeclarationIsNullException(); }

            if (!target.DeclarationType.Equals(DeclarationType.Variable)) { throw new InvalidDeclarationTypeException(target); }

            _targetQMN = target.QualifiedModuleName;

            var builder = new EncapsulateFieldElementsBuilder(_declarationFinderProvider, _targetQMN);

            var selected = builder.Candidates.Single(c => c.Declaration == target);
            selected.EncapsulateFlag = true;

            var model = new EncapsulateFieldModel(
                                target,
                                builder.Candidates,
                                builder.ObjectStateUDTCandidates,
                                builder.DefaultObjectStateUDT,
                                PreviewRewrite,
                                _declarationFinderProvider,
                                builder.ValidationsProvider);

            if (builder.ObjectStateUDT != null)
            {
                model.EncapsulateFieldStrategy = EncapsulateFieldStrategy.ConvertFieldsToUDTMembers;
                model.ObjectStateUDTField = builder.ObjectStateUDT;
            }

            return model;
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            var refactorRewriteSession = new EncapsulateFieldRewriteSession(RewritingManager.CheckOutCodePaneSession()) as IEncapsulateFieldRewriteSession;

            refactorRewriteSession = RefactorRewrite(model, refactorRewriteSession);

            if (!refactorRewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(refactorRewriteSession.RewriteSession);
            }
        }

        private string PreviewRewrite(EncapsulateFieldModel model)
        {
            var previewSession = new EncapsulateFieldRewriteSession(RewritingManager.CheckOutCodePaneSession()) as IEncapsulateFieldRewriteSession; ;

            previewSession = RefactorRewrite(model, previewSession, true);

            return previewSession.CreatePreview(_targetQMN);
        }

        private IEncapsulateFieldRewriteSession RefactorRewrite(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession, bool asPreview = false)
        {
            if (!model.SelectedFieldCandidates.Any()) { return refactorRewriteSession; }

            var strategy = model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                ? new ConvertFieldsToUDTMembers(_declarationFinderProvider, model, _indenter) as IEncapsulateStrategy
                : new UseBackingFields(_declarationFinderProvider, model, _indenter) as IEncapsulateStrategy;

            return strategy.RefactorRewrite(refactorRewriteSession, asPreview);
        }
    }
}
