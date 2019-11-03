using System;
using System.Linq;
using Rubberduck.Common;
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
//<<<<<<< HEAD

        private List<string> _properties;
        private List<string> _newFields;
//        public EncapsulateFieldRefactoring(IDeclarationFinderProvider declarationFinderProvider, IIndenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
//        :base(rewritingManager, selectionService, factory)
//=======
        
        public EncapsulateFieldRefactoring(
            IDeclarationFinderProvider declarationFinderProvider, 
            IIndenter indenter, 
            IRefactoringPresenterFactory factory, 
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
        :base(rewritingManager, selectionProvider, factory)
//>>>>>>> rubberduck-vba/next
        {
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _indenter = indenter;
            _properties = new List<string>();
            _newFields = new List<string>();
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            //<<<<<<< HEAD
            //var selectedDeclaration = _declarationFinderProvider.DeclarationFinder.FindSelectedDeclaration(targetSelection);
            //if (selectedDeclaration.DeclarationType.Equals(DeclarationType.Variable))
            //{
            //    return selectedDeclaration;
            //}

            ////TODO: the code below should no longer be needed after SelectionService PR
            //var variableTarget = _declarationFinderProvider.DeclarationFinder
            //    .UserDeclarations(DeclarationType.Variable)
            //    .FindVariable(targetSelection);

            //if (variableTarget is null)
            //{
            //    //TODO: This is blunt...handles only the simplest cases
            //    //Need the ISelectionService PR
            //    var udtMemberTargets = _declarationFinderProvider.DeclarationFinder
            //    .UserDeclarations(DeclarationType.UserDefinedTypeMember);
            //    return udtMemberTargets.FirstOrDefault();
            //}

            //return variableTarget;
            //=======
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selectedDeclaration == null
                || selectedDeclaration.DeclarationType != DeclarationType.Variable
                || selectedDeclaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return null;
            }

            return selectedDeclaration;
            //>>>>>>> rubberduck-vba/next
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

            var model =  new EncapsulateFieldModel(target);

            var udtVariables = _declarationFinderProvider.DeclarationFinder
                .Members(model.TargetDeclaration.QualifiedName.QualifiedModuleName)
                .Where(v => v.DeclarationType.Equals(DeclarationType.Variable) 
                    && (v.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false));

            var udtVariableRules = new List<EncapsulateUDTVariableRule>();
            foreach (var udtVariable in udtVariables)
            {
                var udtDefinition = GetUDTDefinition(udtVariable);

                var udtVariableRule = new EncapsulateUDTVariableRule()
                {
                    Variable = udtVariable,
                    UserDefinedType = udtDefinition.UserDefinedType,
                    UserDefinedTypeMembers = udtDefinition.UdtMembers,
                    EncapsulateVariable = true,
                    EncapsulateAllUDTMembers = false,
                    UDTMembersToEncapsulate = new List<Declaration>()
                };
                udtVariableRules.Add(udtVariableRule);
            }

            model.CanStoreStateUsingUDTMembers = udtVariableRules.Count() == 1
                || udtVariables.Distinct().Count() == udtVariableRules.Count();

            model.UDTVariableRules = udtVariableRules;

            return model;
        }

        private (Declaration UserDefinedType, IEnumerable<Declaration> UdtMembers) GetUDTDefinition(Declaration udtVariable)
        {
            var udt = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedType)
                .Where(ut => ut.IdentifierName.Equals(udtVariable.AsTypeName))
                .Single();

            var udtMembers = _declarationFinderProvider.DeclarationFinder
               .UserDeclarations(DeclarationType.UserDefinedTypeMember)
               .Where(utm => udt.IdentifierName == utm.ParentDeclaration.IdentifierName);
            return (udt, udtMembers);
        }

        protected override void RefactorImpl(EncapsulateFieldModel modelConcrete)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();

            var model = modelConcrete as IEncapsulateFieldModel;
            foreach (var target in model.EncapsulationTargets)
            {
                UpdateReferences(target, rewriteSession);
                if (target is IEncapsulateUDTMember udtTarget)
                {
                    AddUDTMemberProperty(udtTarget, modelConcrete.TargetDeclaration.QualifiedModuleName);
                }
                else
                {
                    AddProperty(target);
                }
            }

            EnforceEncapsulatedVariablePrivateAccessibility(modelConcrete.TargetDeclaration, rewriteSession);

            InsertProperties(modelConcrete.TargetDeclaration.QualifiedModuleName, rewriteSession);

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private int DeterminePropertyInsertionTokenIndex(QualifiedModuleName qualifiedModuleName)
        {
            return _declarationFinderProvider.DeclarationFinder
                    .Members(qualifiedModuleName)
                    .Where(d => d.DeclarationType == DeclarationType.Variable
                                && !d.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
                    .OrderBy(declaration => declaration.QualifiedSelection)
                    .Last().Context.Stop.TokenIndex;
        }

        private void AddProperty(IEncapsulateField model)
        {
            var generator = new PropertyGenerator
            {
                PropertyName = model.PropertyName,
                AsTypeName = model.TargetDeclaration.AsTypeName,
                BackingField = model.TargetDeclaration.IdentifierName,
                ParameterName = model.ParameterName,
                GenerateSetter = model.ImplementSetSetterType,
                GenerateLetter = model.ImplementLetSetterType
            };

            _properties.Add(GetPropertyText(generator));
        }

        private void AddUDTMemberProperty(IEncapsulateUDTMember udtMember, QualifiedModuleName qmnInstanceModule)
        {
            var udtDeclaration = _declarationFinderProvider.DeclarationFinder
                .Members(qmnInstanceModule) //udtMember.TargetDeclaration.QualifiedModuleName)
                .Where(m => m.DeclarationType.Equals(DeclarationType.Variable)
                    && udtMember.TargetDeclaration.ParentDeclaration.IdentifierName == m.AsTypeName);

            var generator = new PropertyGenerator
            {
                PropertyName = udtMember.TargetDeclaration.IdentifierName,
                AsTypeName = udtMember.TargetDeclaration.AsTypeName,
                BackingField = $"{udtDeclaration.First().IdentifierName}.{udtMember.TargetDeclaration.IdentifierName}",
                ParameterName = udtMember.ParameterName,
                GenerateSetter = udtMember.TargetDeclaration.IsObject, //model.ImplementSetSetterType,
                GenerateLetter = !udtMember.TargetDeclaration.IsObject //model.ImplementLetSetterType
            };


            _properties.Add(GetPropertyText(generator));
        }

        private void InsertProperties(QualifiedModuleName qualifiedModuleName, IRewriteSession rewriteSession)
        {

            var insertionTokenIndex = DeterminePropertyInsertionTokenIndex(qualifiedModuleName);
            var carriageReturns = $"{Environment.NewLine}{Environment.NewLine}";

            var rewriter = rewriteSession.CheckOutModuleRewriter(qualifiedModuleName);
            rewriter.InsertAfter(insertionTokenIndex, $"{carriageReturns}{string.Join(carriageReturns, _properties)}");
        }

        private void EnforceEncapsulatedVariablePrivateAccessibility(Declaration target, IRewriteSession rewriteSession)
        {
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

        private void UpdateReferences(IEncapsulateField model, IRewriteSession rewriteSession)
        {
            foreach (var reference in model.TargetDeclaration.References)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                rewriter.Replace(reference.Context, model.PropertyName);
            }
        }

        private string GetPropertyText(PropertyGenerator generator)
        {
            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines, true));
        }
    }
}
