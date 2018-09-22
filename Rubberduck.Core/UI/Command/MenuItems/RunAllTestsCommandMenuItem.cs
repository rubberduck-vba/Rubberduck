using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RunAllTestsCommandMenuItem : CommandMenuItemBase
    {
        public RunAllTestsCommandMenuItem(RunAllTestsCommand command)
            : base(command)
        {
        }

        public override string Key => "TestMenu_RunAllTests";
        public override int DisplayOrder => (int)UnitTestingMenuItemDisplayOrder.RunAllTests;
        public override Image Image => Resources.CommandBarIcons.AllLoadedTests;
        public override Image Mask => Resources.CommandBarIcons.AllLoadedTestsMask;
    }
}
