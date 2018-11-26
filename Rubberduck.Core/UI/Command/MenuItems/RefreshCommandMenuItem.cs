using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefreshCommandMenuItem : CommandMenuItemBase
    {
        public RefreshCommandMenuItem(RefreshCommand command) : base(command)
        {
        }
        public override string Key => "RubberduckMenu_Refresh";
        public override int DisplayOrder => (int)RubberduckMenuItemDisplayOrder.Refresh;
    }
}
