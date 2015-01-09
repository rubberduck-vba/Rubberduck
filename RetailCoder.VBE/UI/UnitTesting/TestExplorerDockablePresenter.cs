using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.UI.UnitTesting
{
    [ComVisible(false)]
    public class TestExplorerDockablePresenter : DockablePresenterBase
    {
        // todo: move stuff from TestEngine into here.

        private readonly Parser _parser;
        private TestExplorerWindow Control { get { return UserControl as TestExplorerWindow; } }

        public TestExplorerDockablePresenter(Parser parser, VBE vbe, AddIn addin, IDockableUserControl control) 
            : base(vbe, addin, control)
        {
            _parser = parser;
            RegisterTestExplorerEvents();
        }

        private void RegisterTestExplorerEvents()
        {
            Control.OnRefreshListButtonClick += OnExplorerRefreshListButtonClick;

            Control.OnRunAllTestsButtonClick += OnExplorerRunAllTestsButtonClick;
            Control.OnRunFailedTestsButtonClick += OnExplorerRunFailedTestsButtonClick;
            Control.OnRunLastRunTestsButtonClick += OnExplorerRunLastRunTestsButtonClick;
            Control.OnRunNotRunTestsButtonClick += OnExplorerRunNotRunTestsButtonClick;
            Control.OnRunPassedTestsButtonClick += OnExplorerRunPassedTestsButtonClick;
            Control.OnRunSelectedTestButtonClick += OnExplorerRunSelectedTestButtonClick;

            Control.OnGoToSelectedTest += OnExplorerGoToSelectedTest;

            Control.OnAddExpectedErrorTestMethodButtonClick += OnExplorerAddExpectedErrorTestMethodButtonClick;
            Control.OnAddTestMethodButtonClick += OnExplorerAddTestMethodButtonClick;
            Control.OnAddTestModuleButtonClick += OnExplorerAddTestModuleButtonClick;
        }

        private void OnExplorerRefreshListButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerRunAllTestsButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerRunFailedTestsButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerRunLastRunTestsButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerRunNotRunTestsButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerRunPassedTestsButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerRunSelectedTestButtonClick(object sender, SelectedTestEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerGoToSelectedTest(object sender, SelectedTestEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerAddExpectedErrorTestMethodButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerAddTestMethodButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnExplorerAddTestModuleButtonClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
