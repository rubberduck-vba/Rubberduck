using System.Drawing;
using Rubberduck.Properties;
using Rubberduck.UI.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class RunAllTestsCommand : ICommand
    {
        private readonly TestExplorerDockablePresenter _presenter;

        public RunAllTestsCommand(TestExplorerDockablePresenter presenter)
        {
            _presenter = presenter;
        }

        public void Execute()
        {
            _presenter.RunTests();
        }
    }

    public class RunAllTestsUnitTestingCommandMenuItem : CommandMenuItemBase
    {
        public RunAllTestsUnitTestingCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "TestMenu_RunAllTests"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.RunAllTests; } }
        public override Image Image { get { return Resources.AllLoadedTests_8644_24; } }
        public override Image Mask { get { return Resources.AllLoadedTests_8644_24_Mask; } }
    }
}