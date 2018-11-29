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
            var annotations = new List<IAnnotation>();
            var moduleAnnotations = _declarationFinder.FindAnnotations(module).ToList();
            // VBE 1-based indexing
            for (var currentLine = line - 1; currentLine >= 1; currentLine--)
            {
                //Identifier annotation sections end at the first line above without an identifier annotation.
                if (!moduleAnnotations.Any(annotation => annotation.QualifiedSelection.Selection.StartLine <= currentLine
                                                            && annotation.QualifiedSelection.Selection.EndLine >= currentLine
                                                            && annotation.AnnotationType.HasFlag(AnnotationType.IdentifierAnnotation)))
                {
                    break;
                }

                var annotationsStartingOnCurrentLine = moduleAnnotations.Where(a => a.QualifiedSelection.Selection.StartLine == currentLine && a.AnnotationType.HasFlag(AnnotationType.IdentifierAnnotation));

                annotations.AddRange(annotationsStartingOnCurrentLine);
            }
            return annotations;
        }
    }
}
