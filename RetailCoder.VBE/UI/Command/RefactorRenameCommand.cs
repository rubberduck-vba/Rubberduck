using System;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class RefactorRenameCommand : CommandBase
    {
        public override void Execute(object parameter)
        {
            throw new NotImplementedException();
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