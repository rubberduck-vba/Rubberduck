using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @Ignore annotation, used by Rubberduck to filter inspection results at member and local level.
    /// </summary>
    /// <parameter name="Inspections" type="InspectionNames">
    /// This annotation takes a comma-separated list of inspection names as arguments (at least one is required).
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
            : base("Ignore", AnnotationTarget.General, 1, null, new[] { AnnotationArgumentType.Inspection}, true)
        {}
    }
}
