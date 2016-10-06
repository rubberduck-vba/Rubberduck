namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ICommandBarPopup : ICommandBarControl
    {
        ICommandBar CommandBar { get; }
        ICommandBarControls Controls { get; }
    }
}