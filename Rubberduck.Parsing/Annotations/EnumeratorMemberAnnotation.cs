using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_UserMemId</c> attribute value.
    /// </summary>
    public sealed class EnumeratorMemberAnnotation : AnnotationBase
    {
        public EnumeratorMemberAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.Enumerator, qualifiedSelection)
        {
            Description = parameters.FirstOrDefault();
        }

        public string Description { get; }
    }
}