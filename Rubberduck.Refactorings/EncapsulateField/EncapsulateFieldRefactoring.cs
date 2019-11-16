using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Utility;
using Environment = System.Environment;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoring : InteractiveRefactoringBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
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
            ISelectedDeclarationProvider selectedDeclarationProvider)
        :base(rewritingManager, selectionProvider, factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _indenter = indenter;
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

            var encapsulationCandiateFields = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var userDefinedTypeFieldToTypeDeclarationMap = encapsulationCandiateFields
                .Where(v => v.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false)
                .Select(uv => CreateUDTTuple(uv))
                .ToDictionary(key => key.UDTVariable, element => (element.UserDefinedType, element.UDTMembers));

            var model =  new EncapsulateFieldModel(target, encapsulationCandiateFields, userDefinedTypeFieldToTypeDeclarationMap, _indenter);

            return model;
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();

            var nonUdtMemberFields = model.FlaggedEncapsulationFields
                    .Where(encFld => encFld.Declaration.IsVariable());

            var newContent = new EncapsulateFieldNewContent();
            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;
                newContent = ModifyEncapsulatedVariable(nonUdtMemberField, attributes, rewriteSession, newContent);

                RenameReferences(nonUdtMemberField.Declaration, attributes.PropertyName ?? nonUdtMemberField.Declaration.IdentifierName, rewriteSession);
            }

            newContent = LoadNewPropertyContent(model, newContent);

            var moduleMembers = _declarationFinderProvider.DeclarationFinder
                    .Members(_targetQMN).Where(m => m.IsMember());

            int? codeSectionStartIndex
                = moduleMembers.OrderBy(c => c.Selection)
                            .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, _targetQMN);

            rewriter.InsertNewContent(codeSectionStartIndex, newContent);

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private EncapsulateFieldNewContent LoadNewPropertyContent(EncapsulateFieldModel model, EncapsulateFieldNewContent newContent)
        {
            if (!model.FlaggedEncapsulationFields.Any()) { return newContent; }

            newContent.AddCodeBlock($"{string.Join($"{Environment.NewLine}{Environment.NewLine}", model.PropertiesContent)}");
            return newContent;
        }

        private EncapsulateFieldNewContent ModifyEncapsulatedVariable(IEncapsulatedFieldDeclaration target, IFieldEncapsulationAttributes attributes, IRewriteSession rewriteSession, EncapsulateFieldNewContent newContent)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, _targetQMN);

            if (target.Accessibility == Accessibility.Private && attributes.NewFieldName.Equals(target.IdentifierName))
            {
                rewriter.MakeImplicitDeclarationTypeExplicit(target.Declaration);
                return newContent;
            }

            if (target.Declaration.IsDeclaredInList())
            {
                var targetIdentifier = target.Declaration.Context.GetText().Replace(attributes.FieldName, attributes.NewFieldName);
                var newField = target.Declaration.IsTypeSpecified ? $"{Tokens.Private} {targetIdentifier}" : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {target.Declaration.AsTypeName}";

                rewriter.Remove(target.Declaration);

                newContent.AddDeclarationBlock(newField);
            }
            else
            {
                rewriter.Rename(target.Declaration, attributes.NewFieldName);
                rewriter.SetVariableVisiblity(target.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(target.Declaration);
            }
            return newContent;
        }

        private void RenameReferences(Declaration target, string propertyName, IRewriteSession rewriteSession)
        {
            foreach (var reference in target.References)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                rewriter.Replace(reference.Context, propertyName ?? target.IdentifierName);
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
               .Where(utm => userDefinedTypeDeclaration.IdentifierName == utm.ParentDeclaration.IdentifierName);

            return (udtVariable, userDefinedTypeDeclaration, udtMembers);
        }

    }
}
