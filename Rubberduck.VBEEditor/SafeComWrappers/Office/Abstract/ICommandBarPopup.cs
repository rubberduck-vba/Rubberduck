namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    // Abstraction of the CommandBarPopup coclass interface in the interop assemblies for Office.v8 and Office.v12
    public interface ICommandBarPopup : ICommandBarControl
    {
        ICommandBar CommandBar { get; }
        ICommandBarControls Controls { get; }
    }
}