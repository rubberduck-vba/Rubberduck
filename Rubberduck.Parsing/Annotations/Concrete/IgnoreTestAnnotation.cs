namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @IgnoreTest annotation, used for ignoring a particular unit test in a test module.
    /// </summary>
    /// <parameter>
    /// This annotation takes no argument.
    /// </parameter>
    /// <remarks>
    /// Test Explorer will skip tests decorated with this annotation.
    /// </remarks>
    /// <example>
    /// <module name="Tests" type="Standard Module">
    /// <![CDATA[
    /// '@TestModule
    /// Option Explicit
    ///
    /// '...
    /// 
    /// '@IgnoreTest
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
            : base("IgnoreTest", AnnotationTarget.Member, allowedArguments: 1)
        {}
    }
}