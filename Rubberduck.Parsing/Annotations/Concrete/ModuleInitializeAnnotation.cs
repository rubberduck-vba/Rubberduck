namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute before executing the first unit test in a test module.
    /// </summary>
    public sealed class ModuleInitializeAnnotation : AnnotationBase, ITestAnnotation
    {
        public ModuleInitializeAnnotation()
            : base("ModuleInitialize", AnnotationTarget.Member)
        {}
    }
}
