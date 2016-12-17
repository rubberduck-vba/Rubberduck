using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ShowSourceControlPanelCommandMenuItem : CommandMenuItemBase
    {
        public ShowSourceControlPanelCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key { get { return "ToolsMenu_SourceControl"; } }
        public override int DisplayOrder { get { return (int)ToolsMenuItemDisplayOrder.SourceControl; } }
    }
}
