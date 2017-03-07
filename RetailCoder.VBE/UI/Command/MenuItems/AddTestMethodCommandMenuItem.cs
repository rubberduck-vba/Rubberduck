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

        public override string Key { get { return "TestExplorerMenu_AddTestMethod"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.AddTestMethod; } }

        public override Image Image { get { return Resources.flask; } }
        public override Image Mask { get { return Resources.flask_mask; } }
    }
}
