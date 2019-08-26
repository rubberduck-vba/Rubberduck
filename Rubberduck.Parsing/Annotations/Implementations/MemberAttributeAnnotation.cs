using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// This annotation allows the specification of arbitrary VB_Attribute entries for members.
    /// </summary>
    /// <remarks>
    /// It is disjoint from ModuleAttributeAnnotation because of annotation scoping shenanigans.
    /// </remarks>
    // marked as Variable annotation to accomodate annotations of constants
    // FIXME consider whether type hierarchy is sufficient to mark as Attribute annotation
    // FIXME considre whether this annotation (and ModuleAttribute) should be allowed multiple times
    [Annotation("MemberAttribute", AnnotationTarget.Member | AnnotationTarget.Variable)]
    public class MemberAttributeAnnotation : FlexibleAttributeAnnotationBase
    {
        public MemberAttributeAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IReadOnlyList<string> parameters)
        :base(qualifiedSelection, context, parameters)
        {}
    }
}