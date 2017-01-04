using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class TestExplorerCommandMenuItem : CommandMenuItemBase
    {
        public TestExplorerCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key { get { return "TestMenu_TextExplorer"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.TestExplorer; } }
    }
}
