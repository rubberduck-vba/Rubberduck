using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command
{
    public class RefactorRenameCommand : ICommand
    {
        public void Execute()
        {
            throw new System.NotImplementedException();
        }
    }

    public class RefactorRenameCommandMenuItem : CommandMenuItemBase
    {
        public RefactorRenameCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_Rename"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.RenameIdentifier; } }
    }
}