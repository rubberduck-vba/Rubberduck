using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @Interface annotation, marks a class as an abstract interface; Rubberduck can use this valuable metadata in its code analysis.
    /// </summary>
    /// <remarks>
    /// Code Explorer uses an "interface" icon to represent class modules with this annotation.
    /// </remarks>
    /// <example>
    /// <before>
    /// <module name="Something" type="Class Module">
    /// <![CDATA[
    /// '@Interface
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </before>
    /// <after>
    /// <module name="Something" type="Interface Module">
    /// <![CDATA[
    /// '@Interface
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </after>
    /// </example>
    public sealed class InterfaceAnnotation : AnnotationBase
    {
        public InterfaceAnnotation()
            : base("Interface", AnnotationTarget.Module)
        {}

        public override ComponentType? RequiredComponentType => ComponentType.ClassModule;
    }
}