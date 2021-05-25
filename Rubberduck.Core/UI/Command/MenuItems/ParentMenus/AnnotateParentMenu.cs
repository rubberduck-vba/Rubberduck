using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class AnnotateParentMenu : ParentMenuItemBase
    {
        public AnnotateParentMenu(IEnumerable<IMenuItem> items, IUiDispatcher dispatcher)
            : base(dispatcher, "AnnotateMenu", items)
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
