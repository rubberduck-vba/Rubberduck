using System;
using System.Linq;
using Rubberduck.Parsing;
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

            var allMemberFields = _declarationFinderProvider.DeclarationFinder
                .Members(target.QualifiedModuleName)
                .Where(v => v.DeclarationType.Equals(DeclarationType.Variable)
                    && !v.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member));

            var udtFieldToTypeMap = allMemberFields
                .Where(v => v.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false)
                .Select(uv => CreateUDTTuple(uv))
                .ToDictionary(key => key.UDTVariable, element => (element.UserDefinedType, element.UDTMembers));

            var model =  new EncapsulateFieldModel(target, allMemberFields, udtFieldToTypeMap, _indenter);

            return model;
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

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();

            var nonUdtMemberFields = model.FlaggedEncapsulationFields
                    .Where(efd => efd.DeclarationType.Equals(DeclarationType.Variable));

            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;
                EnforceEncapsulatedVariablePrivateAccessibility(nonUdtMemberField.Declaration, attributes, rewriteSession);
                UpdateReferences(nonUdtMemberField.Declaration, rewriteSession, attributes.PropertyName);
            }

            InsertProperties(model, rewriteSession);

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private void InsertProperties(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            if (!model.FlaggedEncapsulationFields.Any()) { return; }

            var qualifiedModuleName = model.FlaggedEncapsulationFields.First().Declaration.QualifiedModuleName;

            var carriageReturns = $"{Environment.NewLine}{Environment.NewLine}";

            var insertionTokenIndex = _declarationFinderProvider.DeclarationFinder
                    .Members(qualifiedModuleName)
                    .Where(d => d.DeclarationType == DeclarationType.Variable
                                && !d.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
                    .OrderBy(declaration => declaration.QualifiedSelection)
                    .Last().Context.Stop.TokenIndex; ;


            var rewriter = rewriteSession.CheckOutModuleRewriter(qualifiedModuleName);
            rewriter.InsertAfter(insertionTokenIndex, $"{carriageReturns}{string.Join(carriageReturns, model.PropertiesContent)}");
        }

        private void EnforceEncapsulatedVariablePrivateAccessibility(Declaration target, IFieldEncapsulationAttributes attributes, IRewriteSession rewriteSession)
        {
            if (!target.DeclarationType.Equals(DeclarationType.Variable))
            {
                return;
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            if (target.Accessibility == Accessibility.Private && attributes.NewFieldName.Equals(target.IdentifierName))
            {
                if (!target.Context.TryGetChildContext<VBAParser.AsTypeClauseContext>(out _))
                {
                    rewriter.InsertAfter(target.Context.Stop.TokenIndex, $" {Tokens.As} {target.AsTypeName}");
                }
                return;
            }

            if (target.Context.TryGetAncestor<VBAParser.VariableListStmtContext>(out var varList)
                && varList.ChildCount > 1)
            {
                rewriter.Remove(target);

                var arrayDeclaration = target.Context.GetText().Split(new string[] { $" {Tokens.As} " }, StringSplitOptions.None);
                var arrayIdentiferAndDimension = arrayDeclaration[0].Replace(attributes.FieldName, attributes.NewFieldName);

                var newField = target.IsArray ? $"{Accessibility.Private} {arrayIdentiferAndDimension} {Tokens.As} {target.AsTypeName}"
                        : $"{Accessibility.Private} {attributes.NewFieldName} As {target.AsTypeName}";

                rewriter.InsertAfter(varList.Stop.TokenIndex, $"{Environment.NewLine}{newField}");
            }
            else
            {
                var identifierContext = target.Context.GetChild<VBAParser.IdentifierContext>();
                var variableStmtContext = target.Context.GetAncestor<VBAParser.VariableStmtContext>();
                var visibilityContext = variableStmtContext.GetChild<VBAParser.VisibilityContext>();

                rewriter.Replace(visibilityContext, Tokens.Private);
                rewriter.Replace(identifierContext, attributes.NewFieldName);

                if (!target.Context.TryGetChildContext<VBAParser.AsTypeClauseContext>(out _))
                {
                    rewriter.InsertAfter(target.Context.Stop.TokenIndex, $" {Tokens.As} {target.AsTypeName}");
                }
            }
        }

        private void UpdateReferences(Declaration target, IRewriteSession rewriteSession, string newName = null)
        {
            foreach (var reference in target.References)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                rewriter.Replace(reference.Context, newName ?? target.IdentifierName);
            }
        }
    }
}
