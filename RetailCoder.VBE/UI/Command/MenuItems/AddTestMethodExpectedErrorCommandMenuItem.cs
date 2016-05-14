using System.Drawing;
using System.Windows.Input;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddTestMethodExpectedErrorCommandMenuItem : CommandMenuItemBase
    {
        public AddTestMethodExpectedErrorCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "TestExplorer_AddExpectedErrorTestMethod"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.AddTestMethodExpectedError; } }

        public override Image Image { get { return Resources.flask_exclamation; } }
        public override Image Mask { get { return Resources.flask_exclamation_mask; } }
    }
}
