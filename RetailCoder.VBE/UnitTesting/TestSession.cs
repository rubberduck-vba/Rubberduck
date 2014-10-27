using Microsoft.Vbe.Interop;
using RetailCoderVBE.Reflection;
using RetailCoderVBE.UnitTesting.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.UnitTesting
{
    internal class TestSession : IDisposable
    {
        public TestSession(VBE vbe)
        {
            _vbe = vbe;
            _allTests = new Dictionary<TestMethod, TestResult>();
            _explorer = new TestExplorerWindow();
            RegisterTestExplorerEvents();

            _engine = new TestEngine(_explorer);
        }

        private readonly VBE _vbe;

        private readonly TestExplorerWindow _explorer;
        private IDictionary<TestMethod,TestResult> _allTests;
        private IEnumerable<TestMethod> _lastRun;
        private readonly TestEngine _engine;

        private DateTime _timestamp;
        public DateTime Timestamp { get { return _timestamp; } }

        private void ShowExplorerWindow()
        {
            _explorer.Refresh(_allTests);
            if (!_explorer.Visible)
            {
                _explorer.Show();
            }
        }

        public void Run(TestMethod test)
        {
            _timestamp = DateTime.Now;
            _explorer.ClearProgress();
            var tests = new Dictionary<TestMethod, TestResult> { { test, null } };

            _explorer.SetPlayList(tests);
            AssignResults(tests.Keys);

            _lastRun = tests.Keys;
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

            _timestamp = DateTime.Now;

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

            _timestamp = DateTime.Now;
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
        
        public void SynchronizeTests()
        {
            var tests = _vbe.VBProjects
                            .Cast<VBProject>()
                            .SelectMany(project => project.TestMethods())
                            .ToDictionary(test => test, test => _allTests.ContainsKey(test) ? _allTests[test] : null);
            _allTests = tests;
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
            var startLine = 1;
            var startColumn = 1;
            var endLine = -1;
            var endColumn = -1;

            var signature = string.Concat("Public Sub ", e.Selection.MethodName, "()");

            var codeModule = _vbe.VBProjects.Cast<VBProject>()
                                 .Single(project => project.Name == e.Selection.ProjectName)
                                 .VBComponents.Cast<VBComponent>()
                                 .Single(component => component.Name == e.Selection.ModuleName)
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
                _timestamp = DateTime.Now;

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
                _timestamp = DateTime.Now;

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
                    _allTests[test] = _engine.Run(test);
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
                _timestamp = DateTime.Now;

                _explorer.SetPlayList(tests.ToDictionary(test => test, test => null as TestResult));

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
            _explorer.Dispose();
        }
    }
}
