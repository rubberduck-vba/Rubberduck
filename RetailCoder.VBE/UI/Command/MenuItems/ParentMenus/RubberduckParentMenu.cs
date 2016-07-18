using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class RubberduckParentMenu : ParentMenuItemBase
    {
        public RubberduckParentMenu(IEnumerable<IMenuItem> items, int beforeIndex) 
            : base("RubberduckMenu", items, beforeIndex)
        {
        }
    }

    public enum RubberduckMenuItemDisplayOrder
    {
        UnitTesting,
        Refactorings,
        Navigate,
        Tools,
        CodeInspections,
        Settings,
        About,
    }
}
