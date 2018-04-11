using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class SmartIndenterParentMenu : ParentMenuItemBase
    {
        public SmartIndenterParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items)
            : base(buttonFactory, "SmartIndenterMenu", items)
        {
        }

        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.Indenter;
    }

    public enum SmartIndenterMenuItemDisplayOrder
    {
        CurrentProcedure,
        CurrentModule,
        CurrentProject,
        NoIndentAnnotation,
    }
}
