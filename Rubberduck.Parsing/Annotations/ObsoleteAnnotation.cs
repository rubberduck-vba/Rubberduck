using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used to mark members as obsolete, so that Rubberduck can warn users whenever they try to use an obsolete member.
    /// </summary>
    public sealed class ObsoleteAnnotation : AnnotationBase
    {
        public ObsoleteAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.Obsolete, qualifiedSelection)
        {
        }
    }
}