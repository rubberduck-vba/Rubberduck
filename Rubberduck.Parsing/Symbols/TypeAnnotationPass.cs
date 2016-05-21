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
            Stopwatch stopwatch = Stopwatch.StartNew();
            foreach (var declaration in _declarationFinder.FindDeclarationsWithNonBaseAsType())
            {
                AnnotateType(declaration);
            }
            stopwatch.Stop();
            Debug.WriteLine("Type annotations completed in {0}ms.", stopwatch.ElapsedMilliseconds);
        }

        private void AnnotateType(Declaration declaration)
        {
            if (declaration.DeclarationType == DeclarationType.ClassModule || declaration.DeclarationType == DeclarationType.UserDefinedType)
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
            else if (!string.IsNullOrWhiteSpace(declaration.AsTypeNameWithoutArrayDesignator) && !Declaration.BASE_TYPES.Contains(declaration.AsTypeNameWithoutArrayDesignator.ToUpperInvariant()))
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
                // TODO: Reference Collector does not add module, find workaround?
                Debug.WriteLine(string.Format("{0}: Type annotation failed for {1} because module parent is missing.", GetType().Name, typeExpression));
                return;
            }
            var expressionContext = _expressionParser.Parse(typeExpression.Trim());
            var boundExpression = _bindingService.ResolveType(module, declaration.ParentDeclaration, expressionContext);
            if (boundExpression != null)
            {
                declaration.AsTypeDeclaration = boundExpression.ReferencedDeclaration;
            }
            else
            {
                // Commented out due to a massive amount of VT_HRESULT messages.
                //Debug.WriteLine(string.Format("{0}: Failed to resolve type {1}.", GetType().Name, typeExpression));
            }
        }
    }
}
