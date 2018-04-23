namespace Rubberduck.VBEditor.SafeComWrappers
{
    // Abstraction of the MsoBarPosition enum in the interop assemblies for Office.v8 and Office.v12
    public enum CommandBarPosition
    {
        Left = 0,
        Top = 1,
        Right = 2,
        Bottom = 3,
        Floating = 4,
        Popup = 5,
        MenuBar = 6,
    }
}