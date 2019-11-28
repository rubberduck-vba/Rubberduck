using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
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
            ISelectedDeclarationProvider selectedDeclarationProvider)
        : base(rewritingManager, selectionProvider, factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _indenter = indenter;
            _validator = new EncapsulateFieldNamesValidator(_declarationFinderProvider, FlaggedFields);
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
        private IEnumerable<IEncapsulateFieldCandidate> FlaggedFields() => Model?.FlaggedEncapsulationFields ?? Enumerable.Empty<IEncapsulateFieldCandidate>();

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            //Model = model;
            var strategy = model.EncapsulationStrategy;

            var rewriteSession = strategy.RefactorRewrite(model, RewritingManager.CheckOutCodePaneSession());
            strategy.InsertNewContent(CodeSectionStartIndex, model, rewriteSession);

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private string PreviewRewrite(EncapsulateFieldModel model)
        {
            var strategy = model.EncapsulationStrategy;
            var scratchPadRewriteSession = strategy.RefactorRewrite(model, RewritingManager.CheckOutCodePaneSession());

            strategy.InsertNewContent(CodeSectionStartIndex, model, scratchPadRewriteSession, true);

            var previewRewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(scratchPadRewriteSession, _targetQMN);
            var preview = previewRewriter.GetText();

            var counter = 0;
            while ( counter++ < 10 && preview.Contains($"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"))
            {
                preview = preview.Replace($"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}", $"{Environment.NewLine}{Environment.NewLine}");
            }

            return preview;
        }

        private int? CodeSectionStartIndex
        {
            get
            {
                var moduleMembers = _declarationFinderProvider.DeclarationFinder
                        .Members(_targetQMN).Where(m => m.IsMember());

                int? codeSectionStartIndex
                    = moduleMembers.OrderBy(c => c.Selection)
                                .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

                return codeSectionStartIndex;
            }
        }
    }
}
