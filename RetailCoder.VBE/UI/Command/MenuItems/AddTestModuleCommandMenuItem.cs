using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddTestModuleCommandMenuItem : CommandMenuItemBase
    {
        public AddTestModuleCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key { get { return "TestExplorerMenu_AddTestModule"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.AddTestModule; } }
        public override bool BeginGroup { get { return true; } }
    }
}
