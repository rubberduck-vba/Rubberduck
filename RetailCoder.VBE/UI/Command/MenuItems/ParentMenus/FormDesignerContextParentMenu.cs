using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class FormDesignerContextParentMenu : ParentMenuItemBase
    {
        public FormDesignerContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup { get { return true; } }
    }

    public class FormDesignerControlContextParentMenu : ParentMenuItemBase
    {
        public FormDesignerControlContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup { get { return true; } }
    }
}
