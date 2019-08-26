using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public interface IAnnotationUpdater
    {
        void AddAnnotation(IRewriteSession rewriteSession, Declaration declaration, AnnotationAttribute newAnnotation, IReadOnlyList<string> values = null);
        void AddAnnotation(IRewriteSession rewriteSession, IdentifierReference reference, AnnotationAttribute newAnnotation, IReadOnlyList<string> values = null);
        void AddAnnotation(IRewriteSession rewriteSession, QualifiedContext context, AnnotationAttribute newAnnotation, IReadOnlyList<string> values = null);
        void RemoveAnnotation(IRewriteSession rewriteSession, IAnnotation annotation);
        void RemoveAnnotations(IRewriteSession rewriteSession, IEnumerable<IAnnotation> annotations);
        void UpdateAnnotation(IRewriteSession rewriteSession, IAnnotation oldAnnotation, AnnotationAttribute newAnnotation, IReadOnlyList<string> newValues = null);
    }
}