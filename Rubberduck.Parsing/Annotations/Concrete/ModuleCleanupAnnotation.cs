using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @ModuleCleanup annotation, marks a procedure that Rubberduck executes after all tests of a module have completed.
    /// </summary>
    /// <example>
    /// <module name="TestModule1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@TestModule
    /// 
    /// Private Assert As Rubberduck.AssertClass
    /// 
    /// '@ModuleCleanup
    /// Private Sub ModuleCleanup()
    ///     Set Assert = Nothing
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class ModuleCleanupAnnotation : AnnotationBase, ITestAnnotation
    {
        public ModuleCleanupAnnotation()
            : base("ModuleCleanup", AnnotationTarget.Member)
        {}

        public override ComponentType? RequiredComponentType => ComponentType.StandardModule;
    }
}
