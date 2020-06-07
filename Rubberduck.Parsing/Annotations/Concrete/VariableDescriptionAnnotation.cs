namespace Rubberduck.Parsing.Annotations
{
    public class VariableDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public VariableDescriptionAnnotation()
            : base("VariableDescription", AnnotationTarget.Variable, "VB_VarDescription")
        {}
    }
}