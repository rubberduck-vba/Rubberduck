using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ProjectExplorerRefactorRenameCommandMenuItem : CommandMenuItemBase
    {
        public ProjectExplorerRefactorRenameCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_Rename";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.RenameIdentifier;
    }
}
