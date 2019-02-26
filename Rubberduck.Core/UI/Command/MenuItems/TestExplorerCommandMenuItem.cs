using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    internal class TestExplorerCommandMenuItem : CommandMenuItemBase
    {
        public TestExplorerCommandMenuItem(TestExplorerCommand command)
            : base(command)
        {
        }

        public override string Key => "TestMenu_TextExplorer";
        public override int DisplayOrder => (int)UnitTestingMenuItemDisplayOrder.TestExplorer;
    }
}
