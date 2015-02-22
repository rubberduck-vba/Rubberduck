using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.UnitTesting;
using Rubberduck.UI.UnitTesting;
using Rubberduck.Extensions;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerDockablePresenter : DockablePresenterBase
    {
        // todo: move stuff from TestEngine into here.

        private readonly IRubberduckParser _parser;
        private TestExplorerWindow Control { get { return UserControl as TestExplorerWindow; } }
        private readonly ITestEngine _testEngine;

        public TestExplorerDockablePresenter(IRubberduckParser parser, VBE vbe, AddIn addin, IDockableUserControl control, ITestEngine testEngine)
            : base(vbe, addin, control)
        {
            _parser = parser;
            _testEngine = testEngine;
            RegisterTestExplorerEvents();
        }

        public void SynchronizeEngineWithIDE()
        {
            try
            {
                _testEngine.AllTests = this.VBE.VBProjects
                                .Cast<VBProject>()
                                .SelectMany(project => project.TestMethods())
                                .ToDictionary(test => test, test => _testEngine.AllTests.ContainsKey(test) ? _testEngine.AllTests[test] : null);

            }
            catch (ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Two or more projects containing test methods have the same name and identically named tests. Please rename one to continue.",
                    "Warning", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
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
