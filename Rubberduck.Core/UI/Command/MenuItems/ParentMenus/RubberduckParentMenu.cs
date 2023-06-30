using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class RubberduckParentMenu : ParentMenuItemBase
    {
        public RubberduckParentMenu(IEnumerable<IMenuItem> items, int beforeIndex, IUiDispatcher dispatcher) 
            : base(dispatcher, "RubberduckMenu", items, beforeIndex)
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
        Windows,
        CodeInspections,
        Settings,
        About,
    }
}
