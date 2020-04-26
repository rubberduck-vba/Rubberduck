namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @TestInitialize annotation, marks a procedure that the unit testing engine executes once before running each of the tests in a module.
    /// </summary>
    /// <parameter>
    /// This annotation takes no argument.
    /// </parameter>
    /// <example>
    /// <module name="TestModule1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@TestModule
    /// 
    /// '...
    /// Private SUT As Class1
    /// 
    /// '@TestInitialize
    /// Private Sub TestInitialize()
    ///     Set SUT = New Class1
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class TestInitializeAnnotation : AnnotationBase, ITestAnnotation
    {
        public TestInitializeAnnotation()
            : base("TestInitialize", AnnotationTarget.Member)
        {}
    }
}
