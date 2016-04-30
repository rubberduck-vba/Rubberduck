using Rubberduck.Parsing.Binding;
using System.Diagnostics;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class TypeAnnotationPass
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly BindingService _bindingService;
        private readonly BoundExpressionVisitor _boundExpressionVisitor;

        public TypeAnnotationPass(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
            _bindingService = new BindingService(
                new DefaultBindingContext(_declarationFinder),
                new TypeBindingContext(_declarationFinder),
                new ProcedurePointerBindingContext(_declarationFinder));
            _boundExpressionVisitor = new BoundExpressionVisitor();
        }

        public void Annotate()
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            foreach (var declaration in _declarationFinder.FindDeclarationsWithNonBaseAsType())
            {
                AnnotateType(declaration);
            }
            stopwatch.Stop();
            Debug.WriteLine("Type annotations completed in {0}ms.", stopwatch.ElapsedMilliseconds);
        }

        private void AnnotateType(Declaration declarationWithAsType)
        {
            string typeExpression;
            if (declarationWithAsType.IsTypeSpecified())
            {
                var typeContext = declarationWithAsType.GetAsTypeContext();
                typeExpression = typeContext.type().complexType().GetText();
            }
            else
            {
                return;
            }
            var module = Declaration.GetModuleParent(declarationWithAsType);
            if (module == null)
            {
                // TODO: Reference Collector does not add module, find workaround?
                return;
            }
            var boundExpression = _bindingService.ResolveType(module, declarationWithAsType.ParentDeclaration, typeExpression);
            if (boundExpression != null)
            {
                declarationWithAsType.AsTypeDeclaration = boundExpression.ReferencedDeclaration;
            }
        }
    }
}