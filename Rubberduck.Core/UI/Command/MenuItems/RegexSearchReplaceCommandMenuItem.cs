using Rubberduck.Parsing.Common;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    [Disabled]
    public class RegexSearchReplaceCommandMenuItem : CommandMenuItemBase
    {
        public RegexSearchReplaceCommandMenuItem(RegexSearchReplaceCommand command)
            : base(command)
        {
        }

        public override string Key => "RubberduckMenu_RegexSearchReplace";
        public override int DisplayOrder => (int)NavigationMenuItemDisplayOrder.RegexSearchReplace;
    }
}
