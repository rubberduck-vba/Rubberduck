using System.Drawing;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddTestMethodCommandMenuItem : CommandMenuItemBase
    {
        public AddTestMethodCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key => "TestExplorerMenu_AddTestMethod";
        public override int DisplayOrder => (int)UnitTestingMenuItemDisplayOrder.AddTestMethod;

        public override Image Image => Resources.flask;
        public override Image Mask => Resources.flask_mask;
    }
}
