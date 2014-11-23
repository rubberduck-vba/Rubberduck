using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Reflection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
    public class TestEngine : IDisposable
    {
        private readonly VBE _vbe;
        private TestExplorerWindow _explorer;
        private Window _hostWindow;

        public TestEngine(VBE vbe, TestExplorerWindow explorer, Window hostWindow)
        {
            _vbe = vbe;
            _allTests = new Dictionary<TestMethod, TestResult>();
            _explorer = explorer;
            _hostWindow = hostWindow;
            RegisterTestExplorerEvents();
        }

        private IDictionary<TestMethod,TestResult> _allTests;
        private IEnumerable<TestMethod> _lastRun;

        private void ShowExplorerWindow()
        {
            _explorer.Refresh(_allTests);
            if (!_explorer.Visible)
            {
                _explorer.Visible = true;
                _explorer.Show();
            }

            _hostWindow.Visible = true;
        }

        public void Run(TestMethod test)
        {
            _explorer.ClearProgress();
            var tests = new Dictionary<TestMethod, TestResult> { { test, null } };

            _explorer.SetPlayList(tests);
            AssignResults(tests.Keys);

            _lastRun = tests.Keys;
            ShowExplorerWindow();
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            _explorer.ClearProgress();
            var methods = tests.ToDictionary(test => test, test => null as TestResult);

            _explorer.SetPlayList(methods);
            AssignResults(methods.Keys);

            _lastRun = methods.Keys;
            ShowExplorerWindow();
        }

        public void Run()
        {
            _explorer.ClearProgress();

            SynchronizeTests();
            _explorer.SetPlayList(_allTests);

            if (!_allTests.Any())
            {
                _explorer.ClearResults();
                _lastRun = null;
                return;
            }

            var tests = _allTests.Keys;
            AssignResults(tests);

            _lastRun = tests;
            ShowExplorerWindow();
        }

        public void ReRun()
        {
            if (_lastRun == null)
            {
                var tests = _allTests.Keys.ToList();
                foreach (var test in tests)
                {
                    _allTests[test] = null;
                }

                return;
            }

            _explorer.ClearProgress();
            _explorer.SetPlayList(_lastRun.ToDictionary(test => test, test => null as TestResult));

            AssignResults(_lastRun);
            ShowExplorerWindow();
        }

        /// <summary>
        /// Gets the tests from the previous run.
        /// </summary>
        /// <param name="outcome"></param>
        /// <returns></returns>
        public IEnumerable<TestMethod> LastRunTests(TestOutcome? outcome = null)
        {
            return _allTests.Where(test => test.Value != null 
                             && test.Value.Outcome == (outcome ?? test.Value.Outcome))
                             .Select(test => test.Key);
        }
        
        /// <summary>
        /// Finds all tests in all opened projects.
        /// </summary>
        public void SynchronizeTests()
        {
            try
            {
                var tests = _vbe.VBProjects
                                .Cast<VBProject>()
                                .SelectMany(project => project.TestMethods())
                                .ToDictionary(test => test, test => _allTests.ContainsKey(test) ? _allTests[test] : null);

                _allTests = tests;
            }
            catch(ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show("Two or more projects containing test methods have the same name and identically named tests. Please rename one to continue.", "Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
        }

        public void ShowExplorer()
        {
            SynchronizeTests();
            ShowExplorerWindow();
        }

        private void RegisterTestExplorerEvents()
        {
            _explorer.OnRefreshListButtonClick += OnExplorerRefreshListButtonClick;

            _explorer.OnRunAllTestsButtonClick += OnExplorerRunAllTestsButtonClick;
            _explorer.OnRunFailedTestsButtonClick += OnExplorerRunFailedTestsButtonClick;
            _explorer.OnRunLastRunTestsButtonClick += OnExplorerRunLastRunTestsButtonClick;
            _explorer.OnRunNotRunTestsButtonClick += OnExplorerRunNotRunTestsButtonClick;
            _explorer.OnRunPassedTestsButtonClick += OnExplorerRunPassedTestsButtonClick;
            _explorer.OnRunSelectedTestButtonClick += OnExplorerRunSelectedTestButtonClick;

            _explorer.OnGoToSelectedTest += OnExplorerGoToSelectedTest;
            
            _explorer.OnAddExpectedErrorTestMethodButtonClick += OnExplorerAddExpectedErrorTestMethodButtonClick;
            _explorer.OnAddTestMethodButtonClick += OnExplorerAddTestMethodButtonClick;
            _explorer.OnAddTestModuleButtonClick += OnExplorerAddTestModuleButtonClick;
        }

        void OnExplorerRunSelectedTestButtonClick(object sender, SelectedTestEventArgs e)
        {
            this.Run(e.Selection);
        }

        void OnExplorerRunPassedTestsButtonClick(object sender, EventArgs e)
        {
            this.RunPassedTests();
        }

        void OnExplorerRunNotRunTestsButtonClick(object sender, EventArgs e)
        {
            this.RunNotRunTests();
        }

        void OnExplorerRunLastRunTestsButtonClick(object sender, EventArgs e)
        {
            this.ReRun();
        }

        void OnExplorerRunFailedTestsButtonClick(object sender, EventArgs e)
        {
            this.RunFailedTests();
        }

        void OnExplorerRunAllTestsButtonClick(object sender, EventArgs e)
        {
            this.Run();
        }

        void OnExplorerAddTestModuleButtonClick(object sender, EventArgs e)
        {
            NewUnitTestModuleCommand.NewUnitTestModule(_vbe);
            SynchronizeTests();
            _explorer.Refresh(_allTests);
        }

        void OnExplorerAddTestMethodButtonClick(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewTestMethod(_vbe);
            SynchronizeTests();
            _explorer.Refresh(_allTests);
        }

        void OnExplorerAddExpectedErrorTestMethodButtonClick(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewExpectedErrorTestMethod(_vbe);
            SynchronizeTests();
            _explorer.Refresh(_allTests);
        }

        void OnExplorerRefreshListButtonClick(object sender, EventArgs e)
        {
            SynchronizeTests();
            _explorer.Refresh(_allTests);
        }

        void OnExplorerGoToSelectedTest(object sender, SelectedTestEventArgs e)
        {
            var selection = e.Selection.FirstOrDefault();
            if (selection == null)
            {
                return;
            }

            var startLine = 1;
            var startColumn = 1;
            var endLine = -1;
            var endColumn = -1;

            var signature = string.Concat("Public Sub ", selection.MethodName, "()");

            var codeModule = _vbe.VBProjects.Cast<VBProject>()
                                 .First(project => project.Name == selection.ProjectName)
                                 .VBComponents.Cast<VBComponent>()
                                 .First(component => component.Name == selection.ModuleName)
                                 .CodeModule;

            if (codeModule.Find(signature, ref startLine, ref startColumn, ref endLine, ref endColumn))
            {
                codeModule.CodePane.SetSelection(startLine, startColumn, endLine, endColumn);
                codeModule.CodePane.Show();
            }
        }

        public IDictionary<TestMethod, TestResult> AllTests { get { return _allTests; } }

        public IEnumerable<TestMethod> PassedTests
        {
            get
            {
                return _allTests.Where(test => test.Value != null && test.Value.Outcome == TestOutcome.Succeeded)
                             .Select(test => test.Key);
            }
        }

        public void RunPassedTests()
        {
            _explorer.ClearProgress();

            var tests = PassedTests;
            _explorer.SetPlayList(tests.ToDictionary(test => test, test => null as TestResult));

            if (tests.Any())
            {
                AssignResults(tests);

                _lastRun = tests;
                ShowExplorerWindow();
            }
            else
            {
                _explorer.ClearResults();
                _lastRun = null;
            }
        }

        public IEnumerable<TestMethod> FailedTests
        {
            get
            {
                return _allTests.Where(test => test.Value != null && test.Value.Outcome == TestOutcome.Failed)
                             .Select(test => test.Key);
            }
        }

        public void RunFailedTests()
        {
            _explorer.ClearProgress();

            var tests = FailedTests;
            _explorer.SetPlayList(tests.ToDictionary(test => test, test => null as TestResult));

            if (tests.Any())
            {
                AssignResults(tests);

                _lastRun = tests;
                ShowExplorerWindow();
            }
            else
            {
                _explorer.ClearResults();
                _lastRun = null;
            }
        }

        private void AssignResults(IEnumerable<TestMethod> testMethods)
        {
            var tests = testMethods.ToList();
            var keys = _allTests.Keys.ToList();
            foreach (var test in keys)
            {
                if (tests.Contains(test))
                {
                    var result = test.Run();
                    _explorer.WriteResult(test, result);
                    _allTests[test] = result;
                }
                else
                {
                    _allTests[test] = null;
                }
            }
        }

        public IEnumerable<TestMethod> NotRunTests 
        {
            get 
            {
                return _allTests.Where(test => test.Value == null)
                                .Select(test => test.Key);
            }
        }

        public void RunNotRunTests()
        {
            _explorer.ClearProgress();

            var tests = NotRunTests.ToList();
            if (tests.Any())
            {
                _explorer.SetPlayList(tests);

                AssignResults(tests);

                _lastRun = tests;
                ShowExplorerWindow();
            }
            else
            {
                _explorer.ClearResults();
                _lastRun = null;
            }
        }

        public void Dispose()
        {
            _explorer.OnRefreshListButtonClick -= OnExplorerRefreshListButtonClick;

            _explorer.OnRunAllTestsButtonClick -= OnExplorerRunAllTestsButtonClick;
            _explorer.OnRunFailedTestsButtonClick -= OnExplorerRunFailedTestsButtonClick;
            _explorer.OnRunLastRunTestsButtonClick -= OnExplorerRunLastRunTestsButtonClick;
            _explorer.OnRunNotRunTestsButtonClick -= OnExplorerRunNotRunTestsButtonClick;
            _explorer.OnRunPassedTestsButtonClick -= OnExplorerRunPassedTestsButtonClick;
            _explorer.OnRunSelectedTestButtonClick -= OnExplorerRunSelectedTestButtonClick;

            _explorer.OnGoToSelectedTest -= OnExplorerGoToSelectedTest;

            _explorer.OnAddExpectedErrorTestMethodButtonClick -= OnExplorerAddExpectedErrorTestMethodButtonClick;
            _explorer.OnAddTestMethodButtonClick -= OnExplorerAddTestMethodButtonClick;
            _explorer.OnAddTestModuleButtonClick -= OnExplorerAddTestModuleButtonClick;

            _explorer.Dispose();
        }
    }
}
