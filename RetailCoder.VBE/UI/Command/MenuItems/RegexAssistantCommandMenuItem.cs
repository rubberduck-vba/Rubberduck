using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    class RegexAssistantCommandMenuItem : CommandMenuItemBase
    {
        public RegexAssistantCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override string Key { get { return "ToolsMenu_RegexAssistant"; } }

        public override int DisplayOrder { get { return (int)ToolsMenuItemDisplayOrder.RegexAssistant; } }
    }
}
