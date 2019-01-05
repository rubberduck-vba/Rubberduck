using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_Description</c> attribute.
    /// </summary>
    public sealed class DescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public DescriptionAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.Description, qualifiedSelection, context, parameters)
        {}
    }
}