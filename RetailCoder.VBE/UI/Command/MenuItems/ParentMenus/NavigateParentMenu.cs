using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class NavigateParentMenu : ParentMenuItemBase
    {
        public NavigateParentMenu(IEnumerable<IMenuItem> items) 
            : base("RubberduckMenu_Navigate", items)
        {
        }

        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.Navigate; } }
    }
}