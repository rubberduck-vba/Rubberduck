namespace Rubberduck.VBEditor.SafeComWrappers
{
    // Abstraction of the MsoButtonState enum in the interop assemblies for Office.v8 and Office.v12
    public enum ButtonState
    {
        Down = -1,
        Up = 0,
        Mixed = 2,
    }
}