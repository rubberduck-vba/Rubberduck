using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RegexSearchReplaceCommandMenuItem : CommandMenuItemBase
    {
        public RegexSearchReplaceCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_RegexSearchReplace"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.RegexSearchReplace; } }
    }
}
