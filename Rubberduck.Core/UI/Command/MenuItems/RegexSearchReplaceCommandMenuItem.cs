using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
#if !DEBUG
    [Parsing.Common.Disabled]
#endif
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
