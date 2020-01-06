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

            var encapsulationCandidateFactory = new EncapsulateFieldElementFactory(_declarationFinderProvider, _targetQMN);

            var validatorProvider = encapsulationCandidateFactory.ValidatorProvider;
            var candidates = encapsulationCandidateFactory.Candidates;
            var defaultObjectStateUDT = encapsulationCandidateFactory.ObjectStateUDT;

            var selected = candidates.Single(c => c.Declaration == target);
            selected.EncapsulateFlag = true;

            if (TryRetrieveExistingObjectStateUDT(target, candidates, out var objectStateUDT))
            {
                objectStateUDT.IsSelected = true;
                defaultObjectStateUDT.IsSelected = false;
            }

            var model = new EncapsulateFieldModel(
                                target,
                                candidates,
                                defaultObjectStateUDT,
                                PreviewRewrite,
                                _declarationFinderProvider,
                                validatorProvider);

            if (objectStateUDT != null)
            {
                model.StateUDTField = objectStateUDT;
                model.EncapsulateFieldStrategy = EncapsulateFieldStrategy.ConvertFieldsToUDTMembers;
            }

            return model;
        }

        //Identify an existing objectStateUDT and make it unavailable for the user to select for encapsulation.
        //This prevents the user from inadvertently nesting a stateUDT within a new stateUDT
        private bool TryRetrieveExistingObjectStateUDT(Declaration target, IEnumerable<IEncapsulateFieldCandidate> candidates, out IObjectStateUDT objectStateUDT)
        {
            objectStateUDT = null;
            //Determination relies on matching the refactoring-generated name and a couple other UDT attributes
            //to determine if an objectStateUDT pre-exists the refactoring.

            //Question: would using an Annotations (like '@IsObjectStateUDT) be better?
            //The logic would then be: if Annotated => it's the one.  else => apply the matching criteria below
            
            //e.g., In cases where the user chooses an existing UDT for the initial encapsulation, the matching 
            //refactoring will not assign the name and the criteria below will fail => so applying an Annotation would
            //make it possible to find again
            var objectStateUDTIdentifier = $"{EncapsulateFieldResources.StateUserDefinedTypeIdentifierPrefix}{target.QualifiedModuleName.ComponentName}";

            var objectStateUDTMatches = candidates.Where(c => c is IUserDefinedTypeCandidate udt
                    && udt.Declaration.HasPrivateAccessibility()
                    && udt.Declaration.AsTypeDeclaration.IdentifierName.StartsWith(objectStateUDTIdentifier, StringComparison.InvariantCultureIgnoreCase))
                    .Select(pm => pm as IUserDefinedTypeCandidate);

            if (objectStateUDTMatches.Count() == 1)
            {
                objectStateUDT = new ObjectStateUDT(objectStateUDTMatches.First()) { IsSelected = true };
            }
            return objectStateUDT != null;
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
            IEncapsulateFieldRewriteSession refactorRewriteSession = new EncapsulateFieldRewriteSession(RewritingManager.CheckOutCodePaneSession());

            refactorRewriteSession = RefactorRewrite(model, refactorRewriteSession, true);

            var previewRewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            return previewRewriter.GetText(maxConsecutiveNewLines: 3);
        }

        private IEncapsulateFieldRewriteSession RefactorRewrite(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession, bool asPreview = false)
        {
            if (!model.SelectedFieldCandidates.Any()) { return refactorRewriteSession; }

            var strategy = model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers // model.ConvertFieldsToUDTMembers
                ? new ConvertFieldsToUDTMembers(_declarationFinderProvider, model, _indenter) as IEncapsulateStrategy
                : new UseBackingFields(_declarationFinderProvider, model, _indenter) as IEncapsulateStrategy;

            return strategy.RefactorRewrite(model, refactorRewriteSession, asPreview);
        }
    }
}
