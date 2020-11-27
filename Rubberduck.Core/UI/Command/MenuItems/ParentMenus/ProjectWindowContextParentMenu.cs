using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class ProjectWindowContextParentMenu : ParentMenuItemBase
    {
        public ProjectWindowContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }

    public enum ProjectExplorerContextMenuItemDisplayOrder
    {
        RenameIdentifier,
        FindSymbol,
        FindAllReferences,
        FindAllImplementations,
        AddRemoveReferences,
        IgnoreProject,
        UnignoreProject
    }
}
