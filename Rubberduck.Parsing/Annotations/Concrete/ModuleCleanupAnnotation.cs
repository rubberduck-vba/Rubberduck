namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute after all unit tests in a test module have executed.
    /// </summary>
    public sealed class ModuleCleanupAnnotation : AnnotationBase, ITestAnnotation
    {
        public ModuleCleanupAnnotation()
            : base("ModuleCleanup", AnnotationTarget.Member)
        {}
    }
}
