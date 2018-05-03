using Rubberduck.Parsing.Common;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
#if !DEBUG
    [Experimental]
#endif
    public class RegexSearchReplaceCommandMenuItem : CommandMenuItemBase
    {
        public RegexSearchReplaceCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key => "RubberduckMenu_RegexSearchReplace";
        public override int DisplayOrder => (int)NavigationMenuItemDisplayOrder.RegexSearchReplace;
    }
}
