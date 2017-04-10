using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_Description</c> attribute.
    /// </summary>
    public sealed class DescriptionAnnotation : AnnotationBase
    {
        public DescriptionAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.Description, qualifiedSelection)
        {
            Description = parameters.FirstOrDefault();
        }

        public string Description { get; }
    }
}