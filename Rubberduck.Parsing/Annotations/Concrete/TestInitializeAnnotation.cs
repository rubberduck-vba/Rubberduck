using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @TestInitialize annotation, marks a procedure that Rubberduck executes once before running each of the tests in a module.
    /// </summary>
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

        public override ComponentType? RequiredComponentType => ComponentType.StandardModule;
    }
}
