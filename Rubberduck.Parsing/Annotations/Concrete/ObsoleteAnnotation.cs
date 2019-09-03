using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used to mark members as obsolete, so that Rubberduck can warn users whenever they try to use an obsolete member.
    /// </summary>
    public sealed class ObsoleteAnnotation : AnnotationBase
    {
        public string ReplacementDocumentation { get; }

        public ObsoleteAnnotation()
            : base("Obsolete", AnnotationTarget.Member | AnnotationTarget.Variable)
        { }

        // FIXME correctly handle the fact that the replacement documentation is only the first parameter!
    }
}