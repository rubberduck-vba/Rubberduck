using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class IdentifierAnnotationService
    {
        private readonly DeclarationFinder _declarationFinder;

        public IdentifierAnnotationService(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public IEnumerable<IAnnotation> FindAnnotations(QualifiedModuleName module, int line)
        {
            return _declarationFinder.FindAnnotations(module, line).Where(annotation => annotation.AnnotationType.HasFlag(AnnotationType.IdentifierAnnotation));
        }
    }
}
