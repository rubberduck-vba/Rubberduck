using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.UnitTesting.ComCommands;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddTestModuleCommandMenuItem : CommandMenuItemBase
    {
        public AddTestModuleCommandMenuItem(AddTestModuleCommand command)
            : base(command)
        {
        }

        public override string Key => "TestExplorerMenu_AddTestModule";
        public override int DisplayOrder => (int)UnitTestingMenuItemDisplayOrder.AddTestModule;
        public override bool BeginGroup => true;
    }
}
