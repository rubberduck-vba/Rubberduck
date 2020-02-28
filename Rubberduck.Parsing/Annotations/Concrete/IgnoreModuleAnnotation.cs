namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// This annotation allows ignoring inspection results of defined inspections for a whole module
    /// </summary>
    public sealed class IgnoreModuleAnnotation : AnnotationBase
    {
        public IgnoreModuleAnnotation()
            : base("IgnoreModule", AnnotationTarget.Module, allowedArguments: null, allowMultiple: true)
        {}
    }
}