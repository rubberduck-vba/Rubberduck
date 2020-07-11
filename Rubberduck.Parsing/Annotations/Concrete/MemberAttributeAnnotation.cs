namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// This annotation allows the specification of arbitrary VB_Attribute entries for members.
    /// </summary>
    /// <remarks>
    /// It is disjoint from ModuleAttributeAnnotation because of annotation scoping shenanigans.
    /// </remarks>
    // marked as Variable annotation to accomodate annotations of constants
    public class MemberAttributeAnnotation : FlexibleAttributeAnnotationBase
    {
        public MemberAttributeAnnotation()
            : base("MemberAttribute", AnnotationTarget.Member | AnnotationTarget.Variable, _argumentTypes, true)
        {}

        private static AnnotationArgumentType[] _argumentTypes = new[]
        {
            AnnotationArgumentType.Attribute,
            AnnotationArgumentType.Text | AnnotationArgumentType.Number | AnnotationArgumentType.Boolean
        };
    }
}