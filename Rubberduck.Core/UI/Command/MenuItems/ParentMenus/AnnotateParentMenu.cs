using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class AnnotateParentMenu : ParentMenuItemBase
    {
        public AnnotateParentMenu(IEnumerable<IMenuItem> items)
            : base("AnnotateMenu", items)
        {
        }

        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.Annotate;
    }

    public enum AnnotateParentMenuItemDisplayOrder
    {
        SelectedDeclaration,
        SelectedModule,
        SelectedMember,
    }
}
