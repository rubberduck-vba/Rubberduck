namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @Interface annotation, marks a class as an abstract interface; Rubberduck can use this valuable metadata in its code analysis.
    /// </summary>
    /// <parameter>
    /// This annotation takes no argument.
    /// </parameter>
    /// <remarks>
    /// Code Explorer uses an "interface" icon to represent class modules with this annotation.
    /// </remarks>
    /// <example>
    /// <module name="Tests" type="Standard Module">
    /// <![CDATA[
    /// '@Interface
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class InterfaceAnnotation : AnnotationBase
    {
        public InterfaceAnnotation()
            : base("Interface", AnnotationTarget.Module)
        {}
    }
}