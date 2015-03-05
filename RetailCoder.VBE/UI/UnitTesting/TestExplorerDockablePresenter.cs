using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerDockablePresenter : DockablePresenterBase
    {
        private TestExplorerWindow Control { get { return UserControl as TestExplorerWindow; } }
        private readonly ITestEngine _testEngine;

        public TestExplorerDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl control, ITestEngine testEngine)
            : base(vbe, addin, control)
        {
            _testEngine = testEngine;
            RegisterTestExplorerEvents();
        }

        public void Synchronize()
        {
            SynchronizeEngineWithIDE();
            Control.Refresh(_testEngine.AllTests);
        }

        public override void Show()
        {
            Synchronize();
            base.Show();
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

        public void RunTests()
        {
            RunTests(_testEngine.AllTests.Keys);
        }

        public void RunTests(IEnumerable<TestMethod> tests)
        {
            Control.ClearResults(); 
            Control.SetPlayList(tests);
            Control.ClearProgress();
            _testEngine.Run(tests);
        }

        private void TestComplete(object sender, TestCompleteEventArgs e)
        {
            Control.WriteResult(e.Test, e.Result);
        }

        private void OnExplorerRefreshListButtonClick(object sender, EventArgs e)
        {
            Synchronize();
        }

        private void OnExplorerRunAllTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.AllTests.Keys);
        }

        private void OnExplorerRunFailedTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.FailedTests());
        }

        private void OnExplorerRunLastRunTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.LastRunTests());
        }

        private void OnExplorerRunNotRunTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.NotRunTests());
        }

        private void OnExplorerRunPassedTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.PassedTests());
        }

        private void OnExplorerRunSelectedTestButtonClick(object sender, SelectedTestEventArgs e)
        {
            RunTests(e.Selection);
        }

        private void OnExplorerGoToSelectedTest(object sender, SelectedTestEventArgs e)
        {
            var controlSelection = e.Selection.FirstOrDefault();
            if (controlSelection == null)
            {
                return;
            }

            var startLine = 1;
            var startColumn = 1;
            var endLine = -1;
            var endColumn = -1;

            var signature = string.Concat("Public Sub ", controlSelection.MethodName, "()");

            var projects = this.VBE.VBProjects.Cast<VBProject>()
                    .Where(project => project.Name == controlSelection.ProjectName
                                && project.VBComponents
                                    .Cast<VBComponent>()
                                    .Any(c => c.Name == controlSelection.ModuleName)
                           );

            var codeModule = projects.First().VBComponents.Cast<VBComponent>()
                                .First(component => component.Name == controlSelection.ModuleName)
                                .CodeModule;

            if (codeModule.Find(signature, ref startLine, ref startColumn, ref endLine, ref endColumn))
            {
                var selection = new Selection(startLine, startColumn, endLine, endColumn);
                codeModule.CodePane.SetSelection(selection);
            }
        }

        private void OnExplorerAddExpectedErrorTestMethodButtonClick(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewExpectedErrorTestMethod(this.VBE);
            Synchronize();
        }

        private void OnExplorerAddTestMethodButtonClick(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewTestMethod(this.VBE);
            Synchronize();
        }

        private void OnExplorerAddTestModuleButtonClick(object sender, EventArgs e)
        {
            NewUnitTestModuleCommand.NewUnitTestModule(this.VBE);
            Synchronize();
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

            _testEngine.TestComplete += TestComplete;
        }
    }
}
