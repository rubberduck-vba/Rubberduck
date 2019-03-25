using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement.CompilationPasses
{
    public sealed class TypeAnnotationPass : ICompilationPass
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly BindingService _bindingService;
        private readonly VBAExpressionParser _expressionParser;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public TypeAnnotationPass(DeclarationFinder declarationFinder, VBAExpressionParser expressionParser)
        {
            _declarationFinder = declarationFinder;
            var typeBindingContext = new TypeBindingContext(_declarationFinder);
            var procedurePointerBindingContext = new ProcedurePointerBindingContext(_declarationFinder);
            _bindingService = new BindingService(
                _declarationFinder,
                new DefaultBindingContext(_declarationFinder, typeBindingContext, procedurePointerBindingContext),
                typeBindingContext,
                procedurePointerBindingContext);
            _expressionParser = expressionParser;
        }

        public void Execute(IReadOnlyCollection<QualifiedModuleName> modules)
        {
            var toDetermineAsTypeDeclaration = _declarationFinder
                                                .FindDeclarationsWithNonBaseAsType()
                                                .Where(decl => decl.AsTypeDeclaration == null 
                                                        || modules.Contains(decl.QualifiedName.QualifiedModuleName));
            foreach (var declaration in toDetermineAsTypeDeclaration)
            {
                AnnotateType(declaration);
            }
        }

        private void AnnotateType(Declaration declaration)
        {
            if (declaration.DeclarationType.HasFlag(DeclarationType.ClassModule) || 
                declaration.DeclarationType == DeclarationType.UserDefinedType || 
                declaration.DeclarationType == DeclarationType.ComAlias)
            {
                declaration.AsTypeDeclaration = declaration;
                return;
            }
            string typeExpression;
            if (declaration.AsTypeContext != null && declaration.AsTypeContext.type().complexType() != null)
            {
                var typeContext = declaration.AsTypeContext;
                typeExpression = typeContext.type().complexType().GetText();
            }
            else if (!string.IsNullOrWhiteSpace(declaration.AsTypeNameWithoutArrayDesignator) && !SymbolList.BaseTypes.Contains(declaration.AsTypeNameWithoutArrayDesignator.ToUpperInvariant()))
            {
                typeExpression = declaration.AsTypeNameWithoutArrayDesignator;
            }
            else
            {
                return;
            }
            var module = Declaration.GetModuleParent(declaration);
            if (module == null)
            {
                Logger.Warn("Type annotation failed for {0} because module parent is missing.", typeExpression);
                return;
            }
            var expressionContext = _expressionParser.Parse(typeExpression.Trim());
            var boundExpression = _bindingService.ResolveType(module, declaration.ParentDeclaration, expressionContext);
            if (boundExpression.Classification != ExpressionClassification.ResolutionFailed)
            {
                declaration.AsTypeDeclaration = boundExpression.ReferencedDeclaration;
            }
            else
            {
                const string IGNORE_THIS = "DISPATCH";
                if (typeExpression != IGNORE_THIS)
                {
                    Logger.Warn("Failed to resolve type {0}", typeExpression);
                }
            }
        }
    }
}
