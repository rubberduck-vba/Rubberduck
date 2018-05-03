using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RunAllTestsCommandMenuItem : CommandMenuItemBase
    {
        public RunAllTestsCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key => "TestMenu_RunAllTests";
        public override int DisplayOrder => (int)UnitTestingMenuItemDisplayOrder.RunAllTests;
        public override Image Image => Resources.AllLoadedTests;
        public override Image Mask => Resources.AllLoadedTestsMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && state.Status >= ParserState.ResolvedDeclarations && state.Status < ParserState.Error;
        }
    }
}
