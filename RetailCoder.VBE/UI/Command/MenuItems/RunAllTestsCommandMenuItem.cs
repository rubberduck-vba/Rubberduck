using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RunAllTestsCommandMenuItem : CommandMenuItemBase
    {
        public RunAllTestsCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "TestMenu_RunAllTests"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.RunAllTests; } }
        public override Image Image { get { return Resources.AllLoadedTests_8644_24; } }
        public override Image Mask { get { return Resources.AllLoadedTests_8644_24_Mask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready ||
                   state.Status == ParserState.Resolving;
        }
    }
}