using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @ModuleInitialize annotation, marks a procedure that Rubberduck executes before running the first test of a module.
    /// </summary>
    /// <example>
    /// <module name="TestModule1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@TestModule
    /// 
    /// Private Assert As Rubberduck.AssertClass
    /// 
    /// '@ModuleInitialize
    /// Private Sub ModuleInitialize()
    ///     Set Assert = New Rubberduck.AssertClass
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class ModuleInitializeAnnotation : AnnotationBase, ITestAnnotation
    {
        public ModuleInitializeAnnotation()
            : base("ModuleInitialize", AnnotationTarget.Member)
        {}

        public override ComponentType? RequiredComponentType => ComponentType.StandardModule;
    }
}
