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

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoring : InteractiveRefactoringBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IIndenter _indenter;
        private QualifiedModuleName _targetQMN;
        private readonly IEncapsulateFieldNamesValidator _validator;

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
            _validator = new EncapsulateFieldNamesValidator(_declarationFinderProvider);
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

            var encapsulationCandidateFields = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var userDefinedTypeFieldToTypeDeclarationMap = encapsulationCandidateFields
                .Where(v => v.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false)
                .Select(uv => CreateUDTTuple(uv))
                .ToDictionary(key => key.UDTVariable, element => (element.UserDefinedType, element.UDTMembers));

            var model = new EncapsulateFieldModel
                                (target,
                                encapsulationCandidateFields,
                                userDefinedTypeFieldToTypeDeclarationMap,
                                _indenter,
                                _validator,
                                PreviewRewrite);

            return model;
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            var rewriteSession = RefactorRewrite(model, RewritingManager.CheckOutCodePaneSession());

            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, _targetQMN);

            rewriter.InsertNewContent(CodeSectionStartIndex, model.NewContent());

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private string PreviewRewrite(EncapsulateFieldModel model)
        {
            var scratchPadRewriteSession = RefactorRewrite(model, RewritingManager.CheckOutCodePaneSession());

            var previewRewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(scratchPadRewriteSession, _targetQMN);

            var newContent = model.NewContent("'<===== No Changes below this line =====>");

            previewRewriter.InsertNewContent(CodeSectionStartIndex, newContent);

            var preview = previewRewriter.GetText();

            return preview;
        }

        private IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            var nonUdtMemberFields = model.FlaggedEncapsulationFields
                    .Where(encFld => encFld.Declaration.IsVariable());

            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;
                ModifyEncapsulatedVariable(nonUdtMemberField, attributes, rewriteSession, model.EncapsulateWithUDT);
                RenameReferences(nonUdtMemberField, attributes.PropertyName ?? nonUdtMemberField.Declaration.IdentifierName, rewriteSession);
            }

            return rewriteSession;
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

        private void ModifyEncapsulatedVariable(IEncapsulatedFieldDeclaration target, IFieldEncapsulationAttributes attributes, IRewriteSession rewriteSession, bool asUDT = false) //, EncapsulateFieldNewContent newContent)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, _targetQMN);

            if (asUDT)
            {
                rewriter.Remove(target.Declaration);
                return;
            }

            if (target.Declaration.Accessibility == Accessibility.Private && attributes.NewFieldName.Equals(target.Declaration.IdentifierName))
            {
                rewriter.MakeImplicitDeclarationTypeExplicit(target.Declaration);
                return;
            }

            if (target.Declaration.IsDeclaredInList())
            {
                rewriter.Remove(target.Declaration);
            }
            else
            {
                rewriter.Rename(target.Declaration, attributes.NewFieldName);
                rewriter.SetVariableVisiblity(target.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(target.Declaration);
            }
            return;
        }

        private void RenameReferences(IEncapsulatedFieldDeclaration efd, string propertyName, IRewriteSession rewriteSession)
        {
            foreach (var reference in efd.Declaration.References)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                rewriter.Replace(reference.Context, propertyName ?? efd.Declaration.IdentifierName);
            }
        }

        private (Declaration UDTVariable, Declaration UserDefinedType, IEnumerable<Declaration> UDTMembers) CreateUDTTuple(Declaration udtVariable)
        {
            var userDefinedTypeDeclaration = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedType)
                .Where(ut => ut.IdentifierName.Equals(udtVariable.AsTypeName)
                    && (ut.Accessibility.Equals(Accessibility.Private)
                            && ut.QualifiedModuleName == udtVariable.QualifiedModuleName)
                    || (ut.Accessibility != Accessibility.Private))
                    .SingleOrDefault();

            var udtMembers = _declarationFinderProvider.DeclarationFinder
               .UserDeclarations(DeclarationType.UserDefinedTypeMember)
               .Where(utm => userDefinedTypeDeclaration.IdentifierName == utm.ParentDeclaration.IdentifierName
                    && utm.QualifiedModuleName == userDefinedTypeDeclaration.QualifiedModuleName);

            return (udtVariable, userDefinedTypeDeclaration, udtMembers);
        }

    }
}
