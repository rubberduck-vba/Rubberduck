using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class UnitTestingParentMenu : ParentMenuItemBase
    {
        public UnitTestingParentMenu(IEnumerable<IMenuItem> items, IUiDispatcher dispatcher) 
            : base(dispatcher, "RubberduckMenu_UnitTests", items)
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
