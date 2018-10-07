using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA.DeclarationCaching;

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
            for (var currentLine = line - 1; currentLine >= 1; currentLine--)
            {
                if (!moduleAnnotations.Any(annotation => annotation.QualifiedSelection.Selection.StartLine <= currentLine
                                                    && annotation.QualifiedSelection.Selection.EndLine >= currentLine))
                {
                    break;
                }

                var annotationsStartingOnCurrentLine = moduleAnnotations.Where(a => a.QualifiedSelection.Selection.StartLine == currentLine);

                annotations.AddRange(annotationsStartingOnCurrentLine);
            }
            return annotations;
        }
    }
}
