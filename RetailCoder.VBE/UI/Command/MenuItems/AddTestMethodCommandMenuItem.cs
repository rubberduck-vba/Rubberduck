using System.Drawing;
using System.Windows.Input;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddTestMethodCommandMenuItem : CommandMenuItemBase
    {
        public AddTestMethodCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "TestExplorer_AddTestMethod"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.AddTestMethod; } }

        public override Image Image { get { return Resources.flask; } }
        public override Image Mask { get { return Resources.flask_mask; } }
    }
}