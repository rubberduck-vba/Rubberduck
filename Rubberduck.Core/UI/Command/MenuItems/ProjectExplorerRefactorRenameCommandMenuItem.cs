using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ProjectExplorerRefactorRenameCommandMenuItem : CommandMenuItemBase
    {
        public ProjectExplorerRefactorRenameCommandMenuItem(ProjectExplorerRefactorRenameCommand command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_Rename";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.RenameIdentifier;
    }
}
