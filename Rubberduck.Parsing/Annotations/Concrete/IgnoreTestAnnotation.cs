using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @IgnoreTest annotation, used by Rubberduck for skipping a particular test when running the tests of a test module.
    /// </summary>
    /// <parameter name="Reason" type="Text">
    /// An optional argument/comment describing the reason for ignoring the test.
    /// </parameter>
    /// <remarks>
    /// Test Explorer will skip tests decorated with this annotation. Use the ignore/un-ignore commands to automatically add or remove this annotation to a particular test without browsing to its source.
    /// </remarks>
    /// <example>
    /// <module name="Tests" type="Standard Module">
    /// <![CDATA[
    /// '@TestModule
    /// Option Explicit
    ///
    /// '...
    /// 
    /// '@IgnoreTest("Foo is currently breaking this test, see issue #42")
    /// '@TestMethod("Category")
    /// Private Sub GivenFoo_DoesBar()
    ///     '...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class IgnoreTestAnnotation : AnnotationBase
    {
        public IgnoreTestAnnotation()
            : base("IgnoreTest", AnnotationTarget.Member, allowedArguments: 1, allowedArgumentTypes: new []{AnnotationArgumentType.Text})
        {}
    }
}