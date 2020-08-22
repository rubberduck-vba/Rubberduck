using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @Folder annotation, used by Rubberduck to represent and organize modules under a custom folder structure.
    /// </summary>
    /// <parameter name="Path" type="Text">
    /// This string literal argument uses the dot "." character to indicate parent/child folders. Consider using folder names that are valid in the file system; PascalCase names are ideal.
    /// </parameter>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@Folder("Parent.Child.SubChild")
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// '@Folder("Parent.Child")
    /// Option Explicit
    ///
    /// Public Sub Macro1()
    ///     With New Class1 '@Folder does not affect namespace or referencing code in any way.
    ///         .DoSomething
    ///     End With
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class FolderAnnotation : AnnotationBase
    {
        public FolderAnnotation()
            : base("Folder", AnnotationTarget.Module, 1, 1, new[] { AnnotationArgumentType.Text})
        {}
    }
}
