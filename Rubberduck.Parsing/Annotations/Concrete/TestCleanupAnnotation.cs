namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute after executing each unit test in a test module.
    /// </summary>
    public sealed class TestCleanupAnnotation : AnnotationBase, ITestAnnotation
    {
        public TestCleanupAnnotation()
            : base("TestCleanup", AnnotationTarget.Member)
        {}
    }
}
