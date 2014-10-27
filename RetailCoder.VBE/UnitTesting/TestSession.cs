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
            _tests = new Dictionary<TestMethod, TestResult>();
            _explorer = new TestExplorerWindow();
            RegisterTestExplorerEvents();

            _engine = new TestEngine(_explorer);
        }

        private readonly VBE _vbe;

        private readonly TestExplorerWindow _explorer;
        private IDictionary<TestMethod,TestResult> _tests;
        private IEnumerable<TestMethod> _lastRun;
        private readonly TestEngine _engine;

        private DateTime _timestamp;
        public DateTime Timestamp { get { return _timestamp; } }

        private void ShowExplorerWindow()
        {
            _explorer.Refresh(_tests);
            if (!_explorer.Visible)
            {
                _explorer.Show();
            }
        }

        public void Run(TestMethod test)
        {
            _timestamp = DateTime.Now;
            _explorer.ClearProgress();
            _tests[test] = _engine.Run(test);
            _explorer.Update();

            _lastRun = new[] { test };
            ShowExplorerWindow();
        }

        public void Run()
        {
            _explorer.ClearProgress();

            SynchronizeTests();
            if (!_tests.Any())
            {
                return;
            }

            _timestamp = DateTime.Now;

            var tests = _tests.Keys;
            AssignResults(tests);

            _lastRun = tests;
            ShowExplorerWindow();
        }

        public void ReRun()
        {
            if (_lastRun == null)
            {
                return;
            }

            _timestamp = DateTime.Now;
            _explorer.ClearProgress();

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
            return _tests.Where(test => test.Value != null 
                             && test.Value.Outcome == (outcome ?? test.Value.Outcome))
                             .Select(test => test.Key);
        }
        
        public void SynchronizeTests()
        {
            var tests = _vbe.VBProjects
                            .Cast<VBProject>()
                            .SelectMany(project => project.TestMethods())
                            .ToDictionary(test => test, test => _tests.ContainsKey(test) ? _tests[test] : null);
            _tests = tests;
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
            _explorer.Refresh(_tests);
        }

        void OnExplorerAddTestMethodButtonClick(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewTestMethod(_vbe);
            SynchronizeTests();
            _explorer.Refresh(_tests);
        }

        void OnExplorerAddExpectedErrorTestMethodButtonClick(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewExpectedErrorTestMethod(_vbe);
            SynchronizeTests();
            _explorer.Refresh(_tests);
        }

        void OnExplorerRefreshListButtonClick(object sender, EventArgs e)
        {
            SynchronizeTests();
            _explorer.Refresh(_tests);
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

        public IDictionary<TestMethod, TestResult> AllTests { get { return _tests; } }

        public IEnumerable<TestMethod> PassedTests
        {
            get
            {
                return _tests.Where(test => test.Value != null && test.Value.Outcome == TestOutcome.Succeeded)
                             .Select(test => test.Key);
            }
        }

        public void RunPassedTests()
        {
            _explorer.ClearProgress();

            var tests = PassedTests;
            if (FailedTests.Any())
            {
                _timestamp = DateTime.Now;

                AssignResults(tests);

                _lastRun = tests;
                ShowExplorerWindow();
            }
        }

        public IEnumerable<TestMethod> FailedTests
        {
            get
            {
                return _tests.Where(test => test.Value != null && test.Value.Outcome == TestOutcome.Failed)
                             .Select(test => test.Key);
            }
        }

        public void RunFailedTests()
        {
            _explorer.ClearProgress();

            var tests = FailedTests;
            if (FailedTests.Any())
            {
                _timestamp = DateTime.Now;

                AssignResults(tests);

                _lastRun = tests;
                ShowExplorerWindow();
            }
        }

        private void AssignResults(IEnumerable<TestMethod> testMethods)
        {
            var tests = testMethods.ToList();
            var keys = _tests.Keys.ToList();
            foreach (var test in keys)
            {
                if (tests.Contains(test))
                {
                    _tests[test] = _engine.Run(test);
                }
                else
                {
                    _tests[test] = TestResult.Unknown();
                }
            }
        }

        public IEnumerable<TestMethod> NotRunTests 
        {
            get 
            {
                return _tests.Where(test => test.Value == null)
                             .Select(test => test.Key);
            }
        }

        public void RunNotRunTests()
        {
            _explorer.ClearProgress();

            if (NotRunTests.Any())
            {
                _timestamp = DateTime.Now;

                var tests = NotRunTests;
                AssignResults(tests);

                _lastRun = tests;
                ShowExplorerWindow();
            }
        }

        public void Dispose()
        {
            _explorer.Dispose();
        }
    }
}
