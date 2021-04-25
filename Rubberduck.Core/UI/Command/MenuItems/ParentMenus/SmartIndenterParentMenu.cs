using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class SmartIndenterParentMenu : ParentMenuItemBase
    {
        public SmartIndenterParentMenu(IEnumerable<IMenuItem> items, IUiDispatcher dispatcher)
            : base(dispatcher, "SmartIndenterMenu", items)
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
