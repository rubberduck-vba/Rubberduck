using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a module's <c>VB_PredeclaredId</c> attribute.
    /// </summary>
    public sealed class PredeclaredIdAnnotation : AnnotationBase, IAttributeAnnotation
    {
        public PredeclaredIdAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.PredeclaredId, qualifiedSelection)
        {
            
        }

        public string Attribute => "VB_PredeclaredId = True";
    }
}