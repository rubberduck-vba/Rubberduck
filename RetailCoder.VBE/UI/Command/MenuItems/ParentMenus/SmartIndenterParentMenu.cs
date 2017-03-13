using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class SmartIndenterParentMenu : ParentMenuItemBase
    {
        public SmartIndenterParentMenu(IEnumerable<IMenuItem> items)
            : base("SmartIndenterMenu", items)
        {
        }

        public override int DisplayOrder
        {
            get { return (int)CodePaneContextMenuItemDisplayOrder.Indenter; }
        }
    }

    public enum SmartIndenterMenuItemDisplayOrder
    {
        CurrentProcedure,
        CurrentModule,
        CurrentProject,
        NoIndentAnnotation,
    }
}
