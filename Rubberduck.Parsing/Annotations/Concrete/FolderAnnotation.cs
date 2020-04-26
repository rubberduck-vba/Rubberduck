namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @Folder annotation, determines where in a custom folder structure a given module appears in the Code Explorer toolwindow.
    /// </summary>
    /// <parameter name="Path" type="String">
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
    /// </example>
    public sealed class FolderAnnotation : AnnotationBase
    {
        public FolderAnnotation()
            : base("Folder", AnnotationTarget.Module, 1, 1)
        {}
    }
}
