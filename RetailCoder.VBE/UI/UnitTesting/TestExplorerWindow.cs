using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public partial class TestExplorerWindow : UserControl, ITestExplorerWindow
    {
        private IList<TestExplorerItem> _playList;
        public DataGridView GridView { get { return testOutputGridView; } }

        private BindingList<TestExplorerItem> _allTests;
        public BindingList<TestExplorerItem> AllTests
        {
            get { return _allTests; }
            set
            {
                _allTests = value;
                testOutputGridView.DataSource = _allTests;
                testOutputGridView.Refresh();
            }
        }

        public string ClassId
        {
            get { return "9CF1392A-2DC9-48A6-AC0B-E601A9802608"; }
        }

        public string Caption
        {
            get { return RubberduckUI.TestExplorerWindow_Caption; }
        }

        public TestExplorerWindow()
        {
            InitializeComponent();

            AllTests = new BindingList<TestExplorerItem>();
            _playList = new List<TestExplorerItem>();

            InitializeGrid();
            RegisterUIEvents();
        }

        private void InitializeGrid()
        {
            testOutputGridView.DataSource = AllTests;
            testOutputGridView.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;

            var messageColumn = testOutputGridView.Columns["Message"];
            if (messageColumn != null)
            {
                messageColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }

            testOutputGridView.Columns["Result"].HeaderText = RubberduckUI.Result;
            testOutputGridView.Columns["QualifiedMemberName"].HeaderText = RubberduckUI.TestExplorer_QualifiedMemberName;
            testOutputGridView.Columns["ProjectName"].HeaderText = RubberduckUI.ProjectName;
            testOutputGridView.Columns["ModuleName"].HeaderText = RubberduckUI.ModuleName;
            testOutputGridView.Columns["MethodName"].HeaderText = RubberduckUI.TestExplorer_MethodName;
            testOutputGridView.Columns["Outcome"].HeaderText = RubberduckUI.Outcome;
            testOutputGridView.Columns["Message"].HeaderText = RubberduckUI.TestExplorer_Message;
            testOutputGridView.Columns["Duration"].HeaderText = RubberduckUI.TestExplorer_Duration;
        }

        private void RegisterUIEvents()
        {
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

            addTestMethodButton.Text = RubberduckUI.TestExplorer_AddTestMethod;
            addTestModuleButton.Text = RubberduckUI.TestExplorer_AddTestModule;
            addExpectedErrorTestMethodButton.Text = RubberduckUI.TestExplorer_AddExpectedErrorTestMethod;
            runAllTestsMenuItem.Text = RubberduckUI.TestExplorer_RunAllTests;
            runFailedTestsMenuItem.Text = RubberduckUI.TestExplorer_RunFailedTests;
            runPassedTestsMenuItem.Text = RubberduckUI.TestExplorer_RunPassedTests;
            runNotRunTestsMenuItem.Text = RubberduckUI.TestExplorer_RunNotRunTests;
            runLastRunMenuItem.Text = RubberduckUI.TestExplorer_RunLastRunTests;
            runSelectedTestMenuItem.Text = RubberduckUI.TestExplorer_RunSelectedTests;
            addButton.Text = RubberduckUI.TestExplorer_AddButtonText;
            runButton.Text = RubberduckUI.TestExplorer_RunButtonText;

            passedTestsLabel.Text = string.Format(RubberduckUI.TestExplorer_TestNumberInconclusive, 0);
            failedTestsLabel.Text = string.Format(RubberduckUI.TestExplorer_TestNumberFailed, 0);
            inconclusiveTestsLabel.Text = string.Format(RubberduckUI.TestExplorer_TestNumberPassed, 0);

            addButton.ToolTipText = RubberduckUI.Add;
            runButton.ToolTipText = RubberduckUI.Run;
            refreshTestsButton.ToolTipText = RubberduckUI.Refresh;
            gotoSelectionButton.ToolTipText = RubberduckUI.TestExplorer_GotoSelectionToolTip;
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
            if (handler != null && AllTests.Any())
            {
                var selection = AllTests.Where(test => testOutputGridView.SelectedRows
                                                                          .Cast<DataGridViewRow>()
                                                                          .Select(row => row.DataBoundItem as TestExplorerItem)
                                                                          .Select(item => item.GetTestMethod())
                                                                          .Contains(test.GetTestMethod()))
                                                                          .ToList(); //ToList forces immediate execution so clearing the gui of previous results won't cause us to lose the selection.

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
            testProgressBar.Maximum = AllTests.Count;
            testProgressBar.Value = 0;
            UpdateCompletedTestsLabels();
        }

        public void ClearResults()
        {
            AllTests = new BindingList<TestExplorerItem>(AllTests.Select(test => new TestExplorerItem(test.GetTestMethod(), null)).ToList());
            testOutputGridView.DataSource = AllTests;
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
            TotalElapsedMilisecondsLabel.Text = string.Format("{0} ms", _playList.Sum(item => item.GetDuration() == TimeSpan.Zero ? 0 : item.GetDuration().Milliseconds));
            passedTestsLabel.Text = string.Format(RubberduckUI.TestExplorer_TestNumberPassed, _playList.Count(item => item.Outcome == TestOutcome.Succeeded.ToString()));
            failedTestsLabel.Text = string.Format(RubberduckUI.TestExplorer_TestNumberFailed, _playList.Count(item => item.Outcome == TestOutcome.Failed.ToString()));
            inconclusiveTestsLabel.Text = string.Format(RubberduckUI.TestExplorer_TestNumberInconclusive, _playList.Count(item => item.Outcome == TestOutcome.Inconclusive.ToString()));
        }

        private TestExplorerItem FindItem(IEnumerable<TestExplorerItem> items, TestMethod test)
        {
            return items.FirstOrDefault(item => item.QualifiedMemberName.Equals(test.QualifiedMemberName));
        }

        public void Refresh(IDictionary<TestMethod, TestResult> tests)
        {
            AllTests = new BindingList<TestExplorerItem>(tests.Select(test => new TestExplorerItem(test.Key, test.Value)).ToList());
            testOutputGridView.DataSource = AllTests;
            testOutputGridView.Refresh();
        }

        public void SetPlayList(IEnumerable<TestMethod> tests)
        {
            SetPlayList(tests.ToDictionary(test => test, test => null as TestResult));
        }

        public void SetPlayList(IDictionary<TestMethod, TestResult> tests)
        {
            _playList = tests.Select(test => new TestExplorerItem(test.Key, test.Value)).ToList();
            UpdateCompletedTestsLabels();
        }

        public event EventHandler OnRefreshListButtonClick;
        private void RefreshTestsButtonClick(object sender, EventArgs e)
        {
            OnButtonClick(OnRefreshListButtonClick);
        }

        public event EventHandler<SelectedTestEventArgs> OnGoToSelectedTest;
        private void GridCellDoubleClicked(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            var handler = OnGoToSelectedTest;
            if (handler != null && e.RowIndex >= 0)
            {
                handler(this, new SelectedTestEventArgs(AllTests[e.RowIndex]));
            }
        }

        private void GotoSelectionButtonClicked(object sender, EventArgs e)
        {
            var handler = OnGoToSelectedTest;
            if (handler != null && AllTests.Any())
            {
                var selectionIndex = testOutputGridView.SelectedRows[0].Index;
                handler(this, new SelectedTestEventArgs(AllTests[selectionIndex]));
            }
        }

        public void WriteResult(TestMethod test, TestResult result)
        {
            var gridItem = FindItem(AllTests, test);
            var playListItem = FindItem(_playList, test);

            if (gridItem == null)
            {
                var item = new TestExplorerItem(test, result);
                AllTests.Add(item);
                gridItem = FindItem(AllTests, test);
            }

            gridItem.SetResult(result);
            playListItem.SetResult(result);

            UpdateProgress();
            testOutputGridView.Refresh();
        }

        public event EventHandler<DataGridViewCellMouseEventArgs> SortColumn;
        private void ColumnHeaderMouseClicked(object sender, DataGridViewCellMouseEventArgs e)
        {
            var handler = SortColumn;
            if (handler == null)
            {
                return;
            }

            handler(this, e);
        }
    }
}
