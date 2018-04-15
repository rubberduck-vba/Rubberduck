using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class UnitTestingParentMenu : ParentMenuItemBase
    {
        public UnitTestingParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items) 
            : base(buttonFactory, "RubberduckMenu_UnitTests", items)
        {
        }

        public override int DisplayOrder => (int)RubberduckMenuItemDisplayOrder.UnitTesting;
    }

    public enum UnitTestingMenuItemDisplayOrder
    {
        TestExplorer,
        RunAllTests,
        AddTestModule,
        AddTestMethod,
        AddTestMethodExpectedError
    }
}
