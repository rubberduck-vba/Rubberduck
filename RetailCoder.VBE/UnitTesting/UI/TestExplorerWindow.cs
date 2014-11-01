using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;

namespace RetailCoderVBE.UnitTesting.UI
{
    internal partial class TestExplorerWindow : Form, ITestOutput
    {
        private BindingList<TestExplorerItem> _allTests;
        private IList<TestExplorerItem> _playList;

        public TestExplorerWindow()
        {
            InitializeComponent();
            InitializeGrid();
            RegisterUIEvents();

            _allTests = new BindingList<TestExplorerItem>();
            _playList = new List<TestExplorerItem>();
        }

        private void InitializeGrid()
        {
            testOutputGridView.DataSource = _allTests;
            testOutputGridView.Columns["Message"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private void RegisterUIEvents()
        {
            FormClosing += TestExplorerWindowFormClosing;

            testOutputGridView.CellDoubleClick += GridCellDoubleClicked;
            testOutputGridView.SelectionChanged += GridSelectionChanged;

            gotoSelectionButton.Click += GotoSelectionButtonClicked;
            addTestMethodButton.Click += AddTestMethodButtonClicked;
            addTestModuleButton.Click += AddTestModuleButtonClicked;
            addExpectedErrorTestMethodButton.Click += AddExpectedErrorTestMethodButtonClicked;
            runAllTestsMenuItem.Click += RunAllTestsMenuItemClicked;
            runFailedTestsMenuItem.Click += RunFailedTestsMenuItemClicked;
            runPassedTestsMenuItem.Click += RunPassedTestsMenuItemClicked;
            runNotRunTestsMenuItem.Click += RunNotRunTestsMenuItemClicked;
            runLastRunMenuItem.Click += RunLastRunMenuItemClicked;
            runSelectedTestMenuItem.Click += RunSelectedTestMenuItemClicked;
        }

        private void GridSelectionChanged(object sender, EventArgs e)
        {
            runSelectedTestMenuItem.Enabled = testOutputGridView.SelectedRows.Count != 0;
        }

        private void OnButtonClick(EventHandler clickEvent)
        {
            var handler = clickEvent;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler<SelectedTestEventArgs> OnRunSelectedTestButtonClick;
        private void RunSelectedTestMenuItemClicked(object sender, EventArgs e)
        {
            var handler = OnRunSelectedTestButtonClick;
            if (handler != null && _allTests.Any())
            {
                var selection = _allTests.Where(test => testOutputGridView.SelectedRows
                                                                          .Cast<DataGridViewRow>()
                                                                          .Select(row => row.DataBoundItem as TestExplorerItem)
                                                                          .Select(item => item.GetTestMethod())
                                                                          .Contains(test.GetTestMethod()));

                handler(this, new SelectedTestEventArgs(selection));
            }
        }

        public event EventHandler OnRunLastRunTestsButtonClick;
        private void RunLastRunMenuItemClicked(object sender, EventArgs e)
        {
            _playList = _playList.Select(test => new TestExplorerItem(test.GetTestMethod(), null)).ToList();
            OnButtonClick(OnRunLastRunTestsButtonClick);
        }

        public event EventHandler OnRunNotRunTestsButtonClick;
        private void RunNotRunTestsMenuItemClicked(object sender, EventArgs e)
        {
            OnButtonClick(OnRunNotRunTestsButtonClick);
        }

        public event EventHandler OnRunPassedTestsButtonClick;
        private void RunPassedTestsMenuItemClicked(object sender, EventArgs e)
        {
            OnButtonClick(OnRunPassedTestsButtonClick);
        }

        public event EventHandler OnRunFailedTestsButtonClick;
        private void RunFailedTestsMenuItemClicked(object sender, EventArgs e)
        {
            OnButtonClick(OnRunFailedTestsButtonClick);
        }

        public event EventHandler OnRunAllTestsButtonClick;
        private void RunAllTestsMenuItemClicked(object sender, EventArgs e)
        {
            OnButtonClick(OnRunAllTestsButtonClick);
        }

        public event EventHandler OnAddExpectedErrorTestMethodButtonClick;
        private void AddExpectedErrorTestMethodButtonClicked(object sender, EventArgs e)
        {
            OnButtonClick(OnAddExpectedErrorTestMethodButtonClick);
        }

        public event EventHandler OnAddTestMethodButtonClick;
        private void AddTestMethodButtonClicked(object sender, EventArgs e)
        {
            OnButtonClick(OnAddTestMethodButtonClick);
        }

        public event EventHandler OnAddTestModuleButtonClick;
        private void AddTestModuleButtonClicked(object sender, EventArgs e)
        {
            OnButtonClick(OnAddTestModuleButtonClick);
        }

        private void TestExplorerWindowFormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }

        public void ClearProgress()
        {
            _completedCount = 0;
            testProgressBar.Maximum = _allTests.Count;
            testProgressBar.Value = 0;
            UpdateCompletedTestsLabels();
        }

        public void ClearResults()
        {
            _allTests = new BindingList<TestExplorerItem>(_allTests.Select(test => new TestExplorerItem(test.GetTestMethod(), null)).ToList());
            testOutputGridView.DataSource = _allTests;
        }

        private int _completedCount; 
        private void UpdateProgress()
        {
            UpdateCompletedTestsLabels();

            runPassedTestsMenuItem.Enabled = _playList.Any(test => test.Outcome == TestOutcome.Succeeded.ToString());
            runFailedTestsMenuItem.Enabled = _playList.Any(test => test.Outcome == TestOutcome.Failed.ToString());

            testProgressBar.Maximum = _playList.Count;
            testProgressBar.Value = ++_completedCount;

            runLastRunMenuItem.Enabled = _completedCount > 0;
        }

        private void UpdateCompletedTestsLabels()
        {
            passedTestsLabel.Text = string.Format("{0} Passed", _playList.Count(item => item.Outcome == TestOutcome.Succeeded.ToString()));
            failedTestsLabel.Text = string.Format("{0} Failed", _playList.Count(item => item.Outcome == TestOutcome.Failed.ToString()));
            inconclusiveTestsLabel.Text = string.Format("{0} Inconclusive", _playList.Count(item => item.Outcome == TestOutcome.Inconclusive.ToString()));
        }

        private TestExplorerItem FindItem(IEnumerable<TestExplorerItem> items, TestMethod test)
        {
            return items.FirstOrDefault(item => item.ProjectName == test.ProjectName
                                                 && item.ModuleName == test.ModuleName
                                                 && item.MethodName == test.MethodName);
        }

        public void Refresh(IDictionary<TestMethod,TestResult> tests)
        {
            _allTests = new BindingList<TestExplorerItem>(tests.Select(test => new TestExplorerItem(test.Key, test.Value)).ToList());
            testOutputGridView.DataSource = _allTests;
            testOutputGridView.Refresh();
        }

        public void SetPlayList(IDictionary<TestMethod,TestResult> tests)
        {
            _playList = tests.Select(test => new TestExplorerItem(test.Key, test.Value)).ToList();
            UpdateCompletedTestsLabels();
        }

        public event EventHandler OnRefreshListButtonClick;
        private void RefreshTestsButtonClick(object sender, System.EventArgs e)
        {
            OnButtonClick(OnRefreshListButtonClick);
        }

        public event EventHandler<SelectedTestEventArgs> OnGoToSelectedTest;
        private void GridCellDoubleClicked(object sender, DataGridViewCellEventArgs e)
        {
            var handler = OnGoToSelectedTest;
            if (handler != null)
            {
                handler(this, new SelectedTestEventArgs(_allTests[e.RowIndex]));
            }
        }

        private void GotoSelectionButtonClicked(object sender, EventArgs e)
        {
            var handler = OnGoToSelectedTest;
            if (handler != null && _allTests.Any())
            {
                var selectionIndex = testOutputGridView.SelectedRows[0].Index;
                handler(this, new SelectedTestEventArgs(_allTests[selectionIndex]));
            }
        }

        public void WriteResult(TestMethod test, TestResult result)
        {
            var gridItem = FindItem(_allTests, test);
            var playListItem = FindItem(_playList, test);

            if (gridItem == null)
            {
                var item = new TestExplorerItem(test, result);
                _allTests.Add(item);
            }
            else
            {
                gridItem.SetResult(result);
                playListItem.SetResult(result);
            }

            UpdateProgress();
            testOutputGridView.Refresh();
        }    
    }

    internal class SelectedTestEventArgs : EventArgs
    {
        public SelectedTestEventArgs(IEnumerable<TestExplorerItem> items)
        {
            _selection = items.Select(item => item.GetTestMethod());
        }

        public SelectedTestEventArgs(TestExplorerItem item)
            : this(new[] { item })
        { }

        private readonly IEnumerable<TestMethod> _selection;
        public IEnumerable<TestMethod> Selection { get { return _selection; } }
    }

    internal class TestExplorerItem
    {
        public TestExplorerItem(TestMethod test, TestResult result)
        {
            _test = test;
            _result = result;
        }

        private readonly TestMethod _test;
        public TestMethod GetTestMethod()
        {
            return _test;
        }

        private TestResult _result;
        public void SetResult(TestResult result)
        {
            _result = result;
        }

        public Image Result { get { return _result.Icon(); } }
        public string ProjectName { get { return _test.ProjectName; } }
        public string ModuleName { get { return _test.ModuleName; } }
        public string MethodName { get { return _test.MethodName; } }
        public string Outcome { get { return _result == null ? string.Empty : _result.Outcome.ToString(); } }
        public string Message { get { return _result == null ? string.Empty : _result.Output; } }
        public string Duration { get { return _result == null ? string.Empty : _result.Duration.ToString() + " ms"; } }
    }

    internal static class TestResultExtensions
    {
        public static Image Icon(this TestResult result)
        {
            var image = RetailCoderVBE.Properties.Resources.question_white;
            if (result != null)
            {
                switch (result.Outcome)
                {
                    case TestOutcome.Succeeded:
                        image = RetailCoderVBE.Properties.Resources.tick_circle;
                        break;

                    case TestOutcome.Failed:
                        image = RetailCoderVBE.Properties.Resources.minus_circle;
                        break;

                    case TestOutcome.Inconclusive:
                        image = RetailCoderVBE.Properties.Resources.exclamation_circle;
                        break;
                }
            }

            return image;
        }
    }
}
