using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @IgnoreModule annotation, used by Rubberduck to filter inspection results module-wide.
    /// </summary>
    /// <parameter name="Inspections" type="InspectionNames">
    /// This annotation optionally takes a comma-separated list of inspection names as argument. If no specific inspection name is provided, then all inspections should ignore the annotated module.
    /// </parameter>
    /// <remarks>
    /// Use this annotation judiciously: while it silences false positives, it also silences legitimate inspection results; useful for muting results in legacy code while still inspecting new code.
    /// </remarks>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@IgnoreModule
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@IgnoreModule UndeclaredVariable, VariableNotUsed, VariableNotAssigned, UnassignedVariableUsage
    ///
    /// Public Sub DoSomething()
    ///     foo = 42
    ///     Debug.Print bar
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class IgnoreModuleAnnotation : AnnotationBase
    {
        public IgnoreModuleAnnotation()
            : base("IgnoreModule", AnnotationTarget.Module, allowedArguments: null, allowedArgumentTypes: new[] { AnnotationArgumentType.Inspection }, allowMultiple: true)
        {}
    }
}