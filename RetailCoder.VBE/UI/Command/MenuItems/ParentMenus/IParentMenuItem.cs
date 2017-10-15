using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public interface IParentMenuItem : IMenuItem, IAppMenu
    {
        ICommandBarControls Parent { get; set; }
        ICommandBarPopup Item { get; }
        void RemoveChildren();
    }
}
