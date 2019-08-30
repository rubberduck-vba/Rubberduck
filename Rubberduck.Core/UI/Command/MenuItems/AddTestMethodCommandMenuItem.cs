using System.Drawing;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.UnitTesting.ComCommands;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddTestMethodCommandMenuItem : CommandMenuItemBase
    {
        public AddTestMethodCommandMenuItem(AddTestMethodCommand command)
            : base(command)
        {
        }

        public override string Key => "TestExplorerMenu_AddTestMethod";
        public override int DisplayOrder => (int)UnitTestingMenuItemDisplayOrder.AddTestMethod;

        public override Image Image => Resources.CommandBarIcons.flask;
        public override Image Mask => Resources.CommandBarIcons.flask_mask;
    }
}
