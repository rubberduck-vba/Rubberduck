using System;
using System.Drawing;
using Rubberduck.Properties;
using Rubberduck.UI.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class TestExplorerCommand : ICommand, IDisposable
    {
        private readonly TestExplorerDockablePresenter _presenter;

        public TestExplorerCommand(TestExplorerDockablePresenter presenter)
        {
            _presenter = presenter;
        }

        public void Execute()
        {
            _presenter.Show();
        }

        public void Dispose()
        {
            _presenter.Dispose();
        }
    }

    public class TestExplorerUnitTestingCommandMenuItem : CommandMenuItemBase
    {
        public TestExplorerUnitTestingCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "TestMenu_TextExplorer"; } }
        public override int DisplayOrder { get { return (int)UnitTestingMenuItemDisplayOrder.TestExplorer; } }
        public override Image Image { get { return Resources.TestManager_8590_32; } }
        public override Image Mask { get { return Resources.TestManager_8590_32_Mask; } }
    }
}