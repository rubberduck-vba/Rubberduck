using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @NoIndent annotation, used by the "indent project" feature to ignore/skip particular modules when bulk-indenting.
    /// </summary>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@NoIndent
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class NoIndentAnnotation : AnnotationBase
    {
        public NoIndentAnnotation()
            : base("NoIndent", AnnotationTarget.Module)
        {}
    }
}
