using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class NavigateParentMenu : ParentMenuItemBase
    {
        public NavigateParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items) 
            : base(buttonFactory, "RubberduckMenu_Navigate", items)
        {
        }

        public override int DisplayOrder => (int)RubberduckMenuItemDisplayOrder.Navigate;
    }

    public enum NavigationMenuItemDisplayOrder
    {
        CodeExplorer,
        RegexSearchReplace,
        FindSymbol,
        FindAllReferences,
        FindImplementations
    }
}
