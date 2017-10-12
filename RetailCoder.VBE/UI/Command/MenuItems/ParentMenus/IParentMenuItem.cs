using Rubberduck.VBEditor.SafeComWrappers.Office.Abstract;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public interface IParentMenuItem : IMenuItem, IAppMenu
    {
        ICommandBarControls Parent { get; set; }
        ICommandBarPopup Item { get; }
        void RemoveMenu();
    }
}
