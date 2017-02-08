using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.VBA;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class TypeAnnotationPass : ICompilationPass
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly BindingService _bindingService;
        private readonly BoundExpressionVisitor _boundExpressionVisitor;
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
            _boundExpressionVisitor = new BoundExpressionVisitor(new AnnotationService(_declarationFinder));
            _expressionParser = expressionParser;
        }

        public void Execute()
        {
            var stopwatch = Stopwatch.StartNew();
            foreach (var declaration in _declarationFinder.FindDeclarationsWithNonBaseAsType())
            {
                AnnotateType(declaration);
            }
            stopwatch.Stop();
        }

        private void AnnotateType(Declaration declaration)
        {
            if (declaration.DeclarationType == DeclarationType.ClassModule || 
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
