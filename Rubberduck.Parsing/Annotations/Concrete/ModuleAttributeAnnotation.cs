namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// This annotation allows specifying arbitrary VB_Attribute entries.
    /// </summary>
    public class ModuleAttributeAnnotation : FlexibleAttributeAnnotationBase
    {
        public ModuleAttributeAnnotation() 
        : base("ModuleAttribute", AnnotationTarget.Module, _argumentTypes, true)
        {}

        private static AnnotationArgumentType[] _argumentTypes = new[]
        {
            AnnotationArgumentType.Attribute,
            AnnotationArgumentType.Text | AnnotationArgumentType.Number | AnnotationArgumentType.Boolean
        };
    }
}