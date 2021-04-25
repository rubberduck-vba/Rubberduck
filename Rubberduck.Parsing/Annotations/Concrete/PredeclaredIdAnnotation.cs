using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @PredeclaredId annotation, indicates the presence of VB_Predeclared module attribute value (True) that defines a default instance for the class, named after that class. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <remarks>
    /// Consider keeping the default/predeclared instance stateless.
    /// </remarks>
    /// <example>
    /// <before>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@PredeclaredId
    /// Option Explicit
    ///
    /// Public Function Create() As Class1
    ///     Set Create = New Class1
    /// End Sub
    /// ]]>
    /// </module>
    /// </before>
    /// <after>
    /// <module name="Class1" type="Predeclared Class">
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = True
    /// '@PredeclaredId
    /// Option Explicit
    ///
    /// Public Function Create() As Class1
    ///     Set Create = New Class1
    /// End Sub
    /// ]]>
    /// </module>
    /// </after>
    /// </example>
    public sealed class PredeclaredIdAnnotation : FixedAttributeValueAnnotationBase
    {
        public PredeclaredIdAnnotation()
            : base("PredeclaredId", AnnotationTarget.Module, "VB_PredeclaredId", new[] { "True" })
        { }

        public override IReadOnlyList<ComponentType> IncompatibleComponentTypes =>
            base.IncompatibleComponentTypes.Concat(new[] { ComponentType.StandardModule }).Distinct().ToList();
    }
}