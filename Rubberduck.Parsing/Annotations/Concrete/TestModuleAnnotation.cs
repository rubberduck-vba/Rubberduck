namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @TestModule annotation, marks a module for unit test discovery.
    /// </summary>
    /// <remarks>
    /// The test engine only scans modules with this annotation when discovering unit tests.
    /// </remarks>
    /// <example>
    /// <module name="TestModule1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@TestModule
    /// 
    /// Private Assert As Rubberduck.AssertClass
    /// '...
    /// ]]>
    /// </module>
    /// </example>
    public sealed class TestModuleAnnotation : AnnotationBase
    {
        public TestModuleAnnotation()
            : base("TestModule", AnnotationTarget.Module)
        {}
    }
}
