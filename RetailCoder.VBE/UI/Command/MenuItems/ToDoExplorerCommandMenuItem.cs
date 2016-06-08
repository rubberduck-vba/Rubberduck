using System.Windows.Input;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ToDoExplorerCommandMenuItem : CommandMenuItemBase
    {
        public ToDoExplorerCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_TodoItems"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.ToDoExplorer; } }
    }
}
