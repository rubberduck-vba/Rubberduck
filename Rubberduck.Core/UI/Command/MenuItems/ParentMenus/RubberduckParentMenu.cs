using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class RubberduckParentMenu : ParentMenuItemBase
    {
        public RubberduckParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items, int beforeIndex) 
            : base(buttonFactory, "RubberduckMenu", items, beforeIndex)
        {
        }
    }

    public enum RubberduckMenuItemDisplayOrder
    {
        Refresh,
        UnitTesting,
        Refactorings,
        Navigate,
        Tools,
        CodeInspections,
        Settings,
        About,
    }
}
