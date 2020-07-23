namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @PredeclaredId annotation, indicates the presence of VB_Predeclared module attribute value (True) that defines a default instance for the class, named after that class. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <remarks>
    /// Consider keeping the default/predeclared instance stateless.
    /// </remarks>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <before>
    /// <![CDATA[
    /// '@PredeclaredId
    /// Option Explicit
    ///
    /// Public Function Create() As Class1
    ///     Set Create = New Class1
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = True
    /// '@PredeclaredId
    /// Option Explicit
    ///
    /// Public Function Create() As Class1
    ///     Set Create = New Class1
    /// End Sub
    /// ]]>
    /// </after>
    /// </module>
    /// </example>
    public sealed class PredeclaredIdAnnotation : FixedAttributeValueAnnotationBase
    {
        public PredeclaredIdAnnotation()
            : base("PredeclaredId", AnnotationTarget.Module, "VB_PredeclaredId", new[] { "True" })
        {}
    }
}