using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @TestModule annotation, marks a module for unit test discovery.
    /// </summary>
    /// <remarks>
    /// Rubberduck only scans modules with this annotation when discovering unit tests in a project.
    /// </remarks>
    /// <example>
    /// <module name="TestModule1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@TestModule
    /// 
    /// Private Assert As Rubberduck.AssertClass
    /// '...
    /// ]]>
    /// </module>
    /// </example>
    public sealed class TestModuleAnnotation : AnnotationBase
    {
        public TestModuleAnnotation()
            : base("TestModule", AnnotationTarget.Module)
        {}

        public override ComponentType? RequiredComponentType => ComponentType.StandardModule;
    }
}
