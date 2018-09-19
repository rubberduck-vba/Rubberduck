using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RegexAssistantCommandMenuItem : CommandMenuItemBase
    {
        public RegexAssistantCommandMenuItem(RegexAssistantCommand command) : base(command)
        {
        }

        public override string Key => "ToolsMenu_RegexAssistant";

        public override int DisplayOrder => (int)ToolsMenuItemDisplayOrder.RegexAssistant;
    }
}
