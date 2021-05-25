using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class FormDesignerContextParentMenu : ParentMenuItemBase
    {
        public FormDesignerContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex, IUiDispatcher dispatcher)
            : base(dispatcher,"RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }

    public class FormDesignerControlContextParentMenu : ParentMenuItemBase
    {
        public FormDesignerControlContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex, IUiDispatcher dispatcher)
            : base(dispatcher,"RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }
}
