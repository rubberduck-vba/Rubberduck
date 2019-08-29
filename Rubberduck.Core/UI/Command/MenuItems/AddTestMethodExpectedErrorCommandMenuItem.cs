using System.Drawing;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.UnitTesting.ComCommands;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddTestMethodExpectedErrorCommandMenuItem : CommandMenuItemBase
    {
        public AddTestMethodExpectedErrorCommandMenuItem(AddTestMethodExpectedErrorCommand command)
            : base(command)
        {
        }

        public override string Key => "TestExplorerMenu_AddExpectedErrorTestMethod";
        public override int DisplayOrder => (int)UnitTestingMenuItemDisplayOrder.AddTestMethodExpectedError;

        public override Image Image => Resources.CommandBarIcons.flask_exclamation;
        public override Image Mask => Resources.CommandBarIcons.flask_exclamation_mask;
    }
}
