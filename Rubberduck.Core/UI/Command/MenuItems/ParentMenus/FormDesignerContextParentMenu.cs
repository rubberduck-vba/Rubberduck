using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class FormDesignerContextParentMenu : ParentMenuItemBase
    {
        public FormDesignerContextParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items, int beforeIndex)
            : base(buttonFactory, "RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }

    public class FormDesignerControlContextParentMenu : ParentMenuItemBase
    {
        public FormDesignerControlContextParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items, int beforeIndex)
            : base(buttonFactory, "RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }
}
