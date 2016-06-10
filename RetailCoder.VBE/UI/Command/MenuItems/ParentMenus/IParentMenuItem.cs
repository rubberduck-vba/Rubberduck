using System;
using Microsoft.Office.Core;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public interface IParentMenuItem : IMenuItem, IAppMenu
    {
        CommandBarControls Parent { get; set; }
        CommandBarPopup Item { get; }
        void RemoveChildren();
    }
}
