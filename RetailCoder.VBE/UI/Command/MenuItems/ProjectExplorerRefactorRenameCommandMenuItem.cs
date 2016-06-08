using System.Windows.Input;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ProjectExplorerRefactorRenameCommandMenuItem : CommandMenuItemBase
    {
        public ProjectExplorerRefactorRenameCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_Rename"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.RenameIdentifier; } }
    }
}
