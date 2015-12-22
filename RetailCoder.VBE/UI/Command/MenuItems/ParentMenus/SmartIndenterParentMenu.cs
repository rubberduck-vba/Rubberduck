using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class SmartIndenterParentMenu : ParentMenuItemBase
    {
        public SmartIndenterParentMenu(IEnumerable<IMenuItem> items)
            : base("SmartIndenterMenu", items)
        {
        }
    }

    public enum SmartIndenterMenuItemDisplayOrder
    {
        CurrentProcedure,
        CurrentModule,
    }
}