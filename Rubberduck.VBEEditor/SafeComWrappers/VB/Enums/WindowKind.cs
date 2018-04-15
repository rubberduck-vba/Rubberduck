namespace Rubberduck.VBEditor.SafeComWrappers
{
    /// <summary>
    /// Values compatible with <see cref="Microsoft.Vbe.Interop.vbext_WindowType"/> enum values.
    /// </summary>
    public enum WindowKind
    {
        CodeWindow = 0,
        Designer = 1,
        Browser = 2,
        Watch = 3,
        Locals = 4,
        Immediate = 5,
        ProjectWindow = 6,
        PropertyWindow = 7,
        Find = 8,
        FindReplace = 9,
        Toolbox = 10,
        LinkedWindowFrame = 11,
        MainWindow = 12,
        ToolWindow = 15,
    }
}