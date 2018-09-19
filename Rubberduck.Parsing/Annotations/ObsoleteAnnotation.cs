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

        public ObsoleteAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.Obsolete, qualifiedSelection, context)
        {
            var firstParameter = parameters.FirstOrDefault();

            ReplacementDocumentation = string.IsNullOrWhiteSpace(firstParameter) ? "" : firstParameter;
        }
    }
}