using System.Collections.Generic;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public interface IAnnotationUpdater
    {
        void AddAnnotation(IRewriteSession rewriteSession, Declaration declaration, IAnnotation newAnnotation, IReadOnlyList<string> values = null);
        void AddAnnotation(IRewriteSession rewriteSession, IdentifierReference reference, IAnnotation newAnnotation, IReadOnlyList<string> values = null);
        void AddAnnotation(IRewriteSession rewriteSession, QualifiedContext context, IAnnotation newAnnotation, IReadOnlyList<string> values = null);
        void RemoveAnnotation(IRewriteSession rewriteSession, IParseTreeAnnotation annotation);
        void RemoveAnnotations(IRewriteSession rewriteSession, IEnumerable<IParseTreeAnnotation> annotations);
        void UpdateAnnotation(IRewriteSession rewriteSession, IParseTreeAnnotation oldAnnotation, IAnnotation newAnnotation, IReadOnlyList<string> newValues = null);
    }
}