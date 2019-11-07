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

            var udtVariables = _declarationFinderProvider.DeclarationFinder
                .Members(target.QualifiedName.QualifiedModuleName)
                .Where(v => v.DeclarationType.Equals(DeclarationType.Variable)
                    && (v.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false));

            var udtTuples = udtVariables.Select(uv => GetUDTDefinition(uv));

            var nonUdtVariables = _declarationFinderProvider.DeclarationFinder
                .Members(target.QualifiedName.QualifiedModuleName)
                .Where(v => v.DeclarationType.Equals(DeclarationType.Variable)
                    && !udtVariables.Contains(v) && !v.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member));

            var model =  new EncapsulateFieldModel(target, nonUdtVariables, udtTuples, _indenter);

            return model;
        }

        private (Declaration UDTVariable, Declaration UserDefinedType, IEnumerable<Declaration> UDTMembers) GetUDTDefinition(Declaration udtVariable)
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

            var rules = new Dictionary<Declaration, EncapsulateUDTVariableRule>();

            foreach (var target in model.EncapsulationTargets)
            {
                if (model.TryGetUDTVariableRule(target.IdentifierName, out EncapsulateUDTVariableRule rule))
                {
                    rules.Add(target, rule);
                }
            }

            var udtVariableTargets = model.EncapsulationTargets.Where(et => rules.ContainsKey(et)).ToList();
            foreach( var udtVariable in udtVariableTargets)
            {
                var rule = rules[udtVariable];
                if (rule.EncapsulateAllUDTMembers)
                {
                    foreach (var udtMember in model.GetUdtMembers(udtVariable))
                    {
                        var udtMemberField = new EncapsulatedValueType(new EncapsulateFieldDeclaration(udtMember));
                        udtMemberField = new EncapsulatedUserDefinedMemberValueType(udtMemberField, new UserDefinedTypeField(udtVariable));
                        udtMemberField.EncapsulationAttributes.Encapsulate = true;
                        model.AddEncapsulationTarget(udtMemberField);
                    }
                }
            }

            foreach (var target in model.EncapsulationTargets)
            {
                if (target.DeclarationType.Equals(DeclarationType.Variable))
                {
                    EnforceEncapsulatedVariablePrivateAccessibility(target, rewriteSession);
                }
                UpdateReferences(target, rewriteSession, model.EncapsulationAttributes(target).PropertyName); // model.PropertyName);
            }

            InsertProperties(model, rewriteSession);

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private void InsertProperties(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            if (!model.EncapsulationTargets.Any()) { return; }

            var qualifiedModuleName = model.EncapsulationTargets.First().QualifiedModuleName;

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

        private void EnforceEncapsulatedVariablePrivateAccessibility(Declaration target, IRewriteSession rewriteSession)
        {
            if (!target.DeclarationType.Equals(DeclarationType.Variable))
            {
                return;
            }

            if (target.Accessibility == Accessibility.Private) { return; }

            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            var newField = $"{Accessibility.Private} {target.IdentifierName} As {target.AsTypeName}";

            if(target.Context.TryGetAncestor<VBAParser.VariableListStmtContext>(out var varList)
                && varList.ChildCount > 1)
            {
                rewriter.Remove(target);
                rewriter.InsertAfter(varList.Stop.TokenIndex, $"{Environment.NewLine}{newField}");
            }
            else
            {
                rewriter.Replace(target.Context.GetAncestor<VBAParser.VariableStmtContext>(), newField);
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
