namespace Rubberduck.VBEditor.SafeComWrappers.Office.Abstract
{
    public interface ICommandBarPopup : ICommandBarControl
    {
        ICommandBar CommandBar { get; }
        ICommandBarControls Controls { get; }
    }
}