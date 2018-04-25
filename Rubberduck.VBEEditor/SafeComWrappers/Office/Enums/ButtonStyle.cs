namespace Rubberduck.VBEditor.SafeComWrappers
{
    // Abstraction of the MsoButtonStyle enum in the interop assemblies for Office.v8 and Office.v12
    public enum ButtonStyle
    {
        Automatic = 0,
        Icon = 1,
        Caption = 2,
        IconAndCaption = 3,
        IconAndWrapCaption = 7,         // Not in Office.v8
        IconAndCaptionBelow = 11,       // Not in Office.v8
        WrapCaption = 14,               // Not in Office.v8
        IconAndWrapCaptionBelow = 15,   // Not in Office.v8
    }
}