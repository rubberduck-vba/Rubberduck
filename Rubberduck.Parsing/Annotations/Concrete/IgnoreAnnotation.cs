namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @Ignore annotation, used for ignoring inspection results at member and local level.
    /// </summary>
    /// <parameter name="Inspections" type="ParamArray (Identifier)">
    /// This annotation optionally takes a comma-separated list of inspection names as argument. If no specific inspection is provided, then all inspections should ignore the annotated target.
    /// </parameter>
    /// <remarks>
    /// Use the @IgnoreModule annotation to annotate at module level.
    /// </remarks>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private InternalState As VBA.Collection
    ///
    /// '@Ignore ProcedureNotUsed
    /// Public Sub DoSomething()
    ///     '@Ignore VariableNotAssigned
    ///     Dim result As Variant
    ///     DoSomething result
    ///     Debug.Print result
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class IgnoreAnnotation : AnnotationBase
    {
        public IgnoreAnnotation()
            : base("Ignore", AnnotationTarget.General, 1, null, true)
        {}
    }
}
