using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class NavigateParentMenu : ParentMenuItemBase
    {
        public NavigateParentMenu(IEnumerable<IMenuItem> items, IUiDispatcher dispatcher) 
            : base(dispatcher,"RubberduckMenu_Navigate", items)
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
