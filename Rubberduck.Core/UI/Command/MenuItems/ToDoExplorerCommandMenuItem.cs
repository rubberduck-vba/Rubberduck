using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ToDoExplorerCommandMenuItem : CommandMenuItemBase
    {
        public ToDoExplorerCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key => "ToolsMenu_TodoItems";
        public override int DisplayOrder => (int)ToolsMenuItemDisplayOrder.ToDoExplorer;
    }
}
