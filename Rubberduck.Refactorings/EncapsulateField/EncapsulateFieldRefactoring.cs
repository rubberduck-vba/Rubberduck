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
        private readonly IIndenter _indenter;

        private List<string> _properties;
        private List<string> _newFields;
        public EncapsulateFieldRefactoring(IDeclarationFinderProvider declarationFinderProvider, IIndenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            _properties = new List<string>();
            _newFields = new List<string>();
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _declarationFinderProvider.DeclarationFinder.FindSelectedDeclaration(targetSelection);
            if (selectedDeclaration.DeclarationType.Equals(DeclarationType.Variable))
            {
                return selectedDeclaration;
            }

            //TODO: the code below should no longer be needed after SelectionService PR
            var variableTarget = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Variable)
                .FindVariable(targetSelection);

            if (variableTarget is null)
            {
                //TODO: This is blunt...handles only the simplest cases
                //Need the ISelectionService PR
                var udtMemberTargets = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedTypeMember);
                return udtMemberTargets.FirstOrDefault();
             }

            return variableTarget;
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

            return new EncapsulateFieldModel(target);
        }

        protected override void RefactorImpl(EncapsulateFieldModel modelConcrete)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();

            var model = modelConcrete as IEncapsulateFieldModel;
            foreach (var target in model.EncapsulationTargets)
            {
                if (target is IEncapsulateUDTMember udtTarget)
                {
                    AddUDTProperties(udtTarget, rewriteSession);
                }
                else
                {
                    AddProperty(target, rewriteSession);
                }
            }

            if (model.EncapsulationTargets.Any(et => et is IEncapsulateUDTMember))
            {
                var members = _declarationFinderProvider.DeclarationFinder
                .Members(modelConcrete.TargetDeclaration.QualifiedName.QualifiedModuleName)
                .OrderBy(declaration => declaration.QualifiedSelection);

                var fields = members.Where(d =>
                                d.DeclarationType == DeclarationType.Variable
                                && !d.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
                                .ToList();

                var rewriter = rewriteSession.CheckOutModuleRewriter(modelConcrete.TargetDeclaration.QualifiedModuleName);
                var carriageReturns = $"{Environment.NewLine}{Environment.NewLine}";
                rewriter.InsertAfter(fields.Last().Context.Stop.TokenIndex, $"{carriageReturns}{string.Join(carriageReturns, _properties)}");
            }


            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private void AddProperty(IEncapsulateField model, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);

            UpdateReferences(model, rewriteSession);

            var members = _declarationFinderProvider.DeclarationFinder
                .Members(model.TargetDeclaration.QualifiedName.QualifiedModuleName)
                .OrderBy(declaration => declaration.QualifiedSelection);

            var insertionPoint = members.Where(d => d.DeclarationType == DeclarationType.Variable && !d.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
                .Last().Context.Stop.TokenIndex;

            _properties.Add(GetPropertyText(model));

            ForceEncapsulatedVariableAccessibility(model.TargetDeclaration, rewriteSession, Accessibility.Private);

            var property = Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine + Environment.NewLine, _properties);
            rewriter.InsertAfter(insertionPoint, property);
        }

        private void AddUDTProperties(IEncapsulateUDTMember udtMember, IRewriteSession rewriteSession)
        {
            var members = _declarationFinderProvider.DeclarationFinder
                .Members(udtMember.TargetDeclaration.QualifiedModuleName)
                .OrderBy(declaration => declaration.QualifiedSelection);

            var fields = members.Where(d => 
                            d.DeclarationType == DeclarationType.Variable 
                            && !d.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
                            .ToList();

            var typeDeclaration = fields.Where(m => m.DeclarationType.Equals(DeclarationType.UserDefinedType)
                && udtMember.TargetDeclaration.ParentDeclaration == m)
                .Single();

            var scopeResolution = members.Where(m => m.DeclarationType.Equals(DeclarationType.Variable)
                && m.AsTypeName.Equals(typeDeclaration.IdentifierName))
                .Select(v => v.IdentifierName).Single();

            var generator = new PropertyGenerator
            {
                PropertyName = udtMember.TargetDeclaration.IdentifierName,
                AsTypeName = udtMember.TargetDeclaration.AsTypeName,
                BackingField = $"{scopeResolution}.{udtMember.TargetDeclaration.IdentifierName}",
                ParameterName = udtMember.ParameterName,
                GenerateSetter = udtMember.TargetDeclaration.IsObject, //model.ImplementSetSetterType,
                GenerateLetter = !udtMember.TargetDeclaration.IsObject //model.ImplementLetSetterType
            };

            UpdateReferences(udtMember, rewriteSession);

            _properties.Add(GetPropertyText(generator));
        }

        private void ForceEncapsulatedVariableAccessibility(Declaration target, IRewriteSession rewriteSession, Accessibility accessibility)
        {
            if (target.Accessibility == accessibility) { return; }

            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            var newField = $"{accessibility} {target.IdentifierName} As {target.AsTypeName}";

            var varList = target.Context.GetAncestor<VBAParser.VariableListStmtContext>();
            if (varList.ChildCount > 1)
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

        private string GetPropertyText(IEncapsulateField model)
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

            return GetPropertyText(generator);
        }

        private string GetPropertyText(PropertyGenerator generator)
        {
            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines, true));
        }
    }
}
