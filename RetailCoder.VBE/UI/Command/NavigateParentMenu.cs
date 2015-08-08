using System.Collections.Generic;

namespace Rubberduck.UI.Command
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