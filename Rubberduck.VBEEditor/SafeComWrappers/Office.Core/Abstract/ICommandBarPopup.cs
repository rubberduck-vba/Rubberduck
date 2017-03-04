namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract
{
    public interface ICommandBarPopup : ICommandBarControl
    {
        ICommandBar CommandBar { get; }
        ICommandBarControls Controls { get; }
    }
}