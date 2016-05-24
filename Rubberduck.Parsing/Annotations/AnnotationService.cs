using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class AnnotationService
    {
        private readonly DeclarationFinder _declarationFinder;

        public AnnotationService(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public IEnumerable<IAnnotation> FindAnnotations(QualifiedModuleName module, int line)
        {
            var annotations = new List<IAnnotation>();
            var moduleAnnotations = _declarationFinder.FindAnnotations(module).ToList();
            // VBE 1-based indexing
            for (var i = line - 1; i >= 1; i--)
            {
                var annotation = moduleAnnotations.SingleOrDefault(a => a.QualifiedSelection.Selection.StartLine == i);
                if (annotation == null)
                {
                    break;
                }
                annotations.Add(annotation);
            }
            return annotations;
        }
    }
}
