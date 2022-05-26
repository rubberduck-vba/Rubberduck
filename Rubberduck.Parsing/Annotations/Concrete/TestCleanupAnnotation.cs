using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @TestCleanup annotation, marks a procedure that Rubberduck executes once after running each of the tests in a module.
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
    /// '@TestCleanup
    /// Private Sub TestCleanup()
    ///     Set SUT = Nothing
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class TestCleanupAnnotation : AnnotationBase, ITestAnnotation
    {
        public TestCleanupAnnotation()
            : base("TestCleanup", AnnotationTarget.Member)
        {}

        public override ComponentType? RequiredComponentType => ComponentType.StandardModule;
    }
}
