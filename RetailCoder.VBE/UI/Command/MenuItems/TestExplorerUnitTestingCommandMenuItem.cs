using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class TestExplorerUnitTestingCommandMenuItem : CommandMenuItemBase
    {
        public TestExplorerUnitTestingCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "TestMenu_TextExplorer"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.TestExplorer; } }
    }
}