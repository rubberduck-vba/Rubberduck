using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for labeling an enum as a bitwise enum
    /// </summary>
    public sealed class FlagsAnnotation : AnnotationBase
    {
        public FlagsAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.Flags, qualifiedSelection, context)
        {
        }
    }
}
