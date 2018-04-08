using System.Drawing;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddTestMethodExpectedErrorCommandMenuItem : CommandMenuItemBase
    {
        public AddTestMethodExpectedErrorCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key => "TestExplorerMenu_AddExpectedErrorTestMethod";
        public override int DisplayOrder => (int)UnitTestingMenuItemDisplayOrder.AddTestMethodExpectedError;

        public override Image Image => Resources.flask_exclamation;
        public override Image Mask => Resources.flask_exclamation_mask;
    }
}
