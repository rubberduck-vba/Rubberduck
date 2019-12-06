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
using Rubberduck.Parsing;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoring : InteractiveRefactoringBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IIndenter _indenter;
        private QualifiedModuleName _targetQMN;
        private IEncapsulateFieldNamesValidator _validator;
        private EncapsulateFieldModel Model { set; get; }

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
            _validator = new EncapsulateFieldNamesValidator(_declarationFinderProvider, CandidateFields);
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
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!target.DeclarationType.Equals(DeclarationType.Variable))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            _targetQMN = target.QualifiedModuleName;

            Model = new EncapsulateFieldModel(
                                target,
                                _declarationFinderProvider,
                                _indenter,
                                _validator,
                                PreviewRewrite);

            return Model;
        }

        //Get rid of Model property after improving the validator ctor
        private IEnumerable<IEncapsulateFieldCandidate> CandidateFields() => Model?.FieldCandidates ?? Enumerable.Empty<IEncapsulateFieldCandidate>();

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            var strategy = model.EncapsulationStrategy;

            var rewriteSession = strategy.RefactorRewrite(model, RewritingManager.CheckOutCodePaneSession());

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private string PreviewRewrite(EncapsulateFieldModel model)
        {
            var strategy = model.EncapsulationStrategy;

            var scratchPadRewriteSession = strategy.GeneratePreview(model, RewritingManager.CheckOutCodePaneSession());

            var previewRewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(scratchPadRewriteSession, _targetQMN);

            return previewRewriter.GetText(maxConsecutiveNewLines: 3);
        }
    }
}
