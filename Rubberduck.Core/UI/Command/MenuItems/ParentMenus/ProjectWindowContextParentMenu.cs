using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class ProjectWindowContextParentMenu : ParentMenuItemBase
    {
        public ProjectWindowContextParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items, int beforeIndex)
            : base(buttonFactory, "RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }
}
