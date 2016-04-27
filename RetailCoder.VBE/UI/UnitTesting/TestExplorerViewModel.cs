using System;
using System.Linq;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;
using Rubberduck.UnitTesting;
using resx = Rubberduck.UI.RubberduckUI;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerViewModel : ViewModelBase, INavigateSelection
    {
        private readonly ITestEngine _testEngine;
        private readonly TestExplorerModelBase _model;
        private readonly IClipboardWriter _clipboard;

        public TestExplorerViewModel(VBE vbe, ITestEngine testEngine, TestExplorerModelBase model, IClipboardWriter clipboard, NewUnitTestModuleCommand newTestModuleCommand, NewTestMethodCommand newTestMethodCommand)
        {
            _testEngine = testEngine;
            _testEngine.TestCompleted += TestEngineTestCompleted;
            _model = model;
            _clipboard = clipboard;

            _navigateCommand = new NavigateCommand();

            _runAllTestsCommand = new RunAllTestsCommand(testEngine, model);
            _addTestModuleCommand = new AddTestModuleCommand(vbe, newTestModuleCommand);
            _addTestMethodCommand = new AddTestMethodCommand(model, newTestMethodCommand);
            _addErrorTestMethodCommand = new AddTestMethodExpectedErrorCommand(model, newTestMethodCommand);

            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand, CanExecuteRefreshCommand);
            _repeatLastRunCommand = new DelegateCommand(ExecuteRepeatLastRunCommand, CanExecuteRepeatLastRunCommand);
            _runNotExecutedTestsCommand = new DelegateCommand(ExecuteRunNotExecutedTestsCommand, CanExecuteRunNotExecutedTestsCommand);
            _runFailedTestsCommand = new DelegateCommand(ExecuteRunFailedTestsCommand, CanExecuteRunFailedTestsCommand);
            _runPassedTestsCommand = new DelegateCommand(ExecuteRunPassedTestsCommand, CanExecuteRunPassedTestsCommand);
            _runSelectedTestCommand = new DelegateCommand(ExecuteSelectedTestCommand, CanExecuteSelectedTestCommand);

            _copyResultsCommand = new DelegateCommand(ExecuteCopyResultsCommand);


        }

        private bool CanExecuteRunPassedTestsCommand(object obj)
        {
            return true; //_model.Tests.Any(test => test.Outcome == TestOutcome.Succeeded);
        }

        private bool CanExecuteRunFailedTestsCommand(object obj)
        {
            return true; //_model.Tests.Any(test => test.Outcome == TestOutcome.Failed);
        }

        private bool CanExecuteRunNotExecutedTestsCommand(object obj)
        {
            return true; //_model.Tests.Any(test => test.Outcome == TestOutcome.Unknown);
        }

        private bool CanExecuteRepeatLastRunCommand(object obj)
        {
            return true; //_model.LastRun.Any();
        }

        public event EventHandler<EventArgs> TestCompleted;
        private void TestEngineTestCompleted(object sender, EventArgs e)
        {
            var handler = TestCompleted;
            if (handler != null)
            {
                handler.Invoke(sender, e);
            }
        }

        public INavigateSource SelectedItem { get { return SelectedTest; } set { SelectedTest = value as TestMethod; } }

        private TestMethod _selectedTest;
        public TestMethod SelectedTest
        {
            get { return _selectedTest; }
            set
            {
                _selectedTest = value;
                OnPropertyChanged();
            }
        }

        private readonly ICommand _runAllTestsCommand;
        public ICommand RunAllTestsCommand { get { return _runAllTestsCommand; } }

        private readonly ICommand _addTestModuleCommand;
        public ICommand AddTestModuleCommand { get { return _addTestModuleCommand; } }

        private readonly ICommand _addTestMethodCommand;
        public ICommand AddTestMethodCommand { get { return _addTestMethodCommand; } }

        private readonly ICommand _addErrorTestMethodCommand;
        public ICommand AddErrorTestMethodCommand { get { return _addErrorTestMethodCommand; } }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand { get { return _refreshCommand; } }

        private readonly ICommand _repeatLastRunCommand;
        public ICommand RepeatLastRunCommand { get { return _repeatLastRunCommand; } }

        private readonly ICommand _runNotExecutedTestsCommand;
        public ICommand RunNotExecutedTestsCommand { get { return _runNotExecutedTestsCommand; } }

        private readonly ICommand _runFailedTestsCommand;
        public ICommand RunFailedTestsCommand { get { return _runFailedTestsCommand; } }

        private readonly ICommand _runPassedTestsCommand;
        public ICommand RunPassedTestsCommand { get { return _runPassedTestsCommand; } }

        private readonly ICommand _copyResultsCommand;
        public ICommand CopyResultsCommand { get { return _copyResultsCommand; } }

        private readonly NavigateCommand _navigateCommand;
        public INavigateCommand NavigateCommand { get { return _navigateCommand; } }

        private readonly ICommand _runSelectedTestCommand;
        public ICommand RunSelectedTestCommand { get { return _runSelectedTestCommand; } }

        public TestExplorerModelBase Model { get { return _model; } }

        private void ExecuteRefreshCommand(object parameter)
        {
            if (Model.IsBusy)
            {
                return;
            }

            _model.Refresh();
            SelectedTest = null;
        }

        private bool CanExecuteRefreshCommand(object parameter)
        {
            return !Model.IsBusy;
        }

        private void ExecuteRepeatLastRunCommand(object parameter)
        {
            var tests = _model.LastRun.ToList();
            _model.ClearLastRun();

            Model.IsBusy = true;
            _testEngine.Run(tests);
            Model.IsBusy = false;
        }

        private void ExecuteRunNotExecutedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            Model.IsBusy = true;
            _testEngine.Run(_model.LastRun.Where(test => test.Result.Outcome == TestOutcome.Unknown));
            Model.IsBusy = false;
        }

        private void ExecuteRunFailedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            Model.IsBusy = true;
            _testEngine.Run(_model.LastRun.Where(test => test.Result.Outcome == TestOutcome.Failed));
            Model.IsBusy = false;
        }

        private void ExecuteRunPassedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            Model.IsBusy = true;
            _testEngine.Run(_model.LastRun.Where(test => test.Result.Outcome == TestOutcome.Succeeded));
            Model.IsBusy = false;
        }

        private bool CanExecuteSelectedTestCommand(object obj)
        {
            return !Model.IsBusy; //true; //SelectedItem != null;
        }

        private void ExecuteSelectedTestCommand(object obj)
        {
            if (SelectedTest == null)
            {
                return;
            }

            _model.ClearLastRun();

            Model.IsBusy = true;
            _testEngine.Run(new[] { SelectedTest });
            Model.IsBusy = false;
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            var results = string.Join("\n", _model.LastRun.Select(test => test.ToString()));
            var passed = _model.LastRun.Count(test => test.Result.Outcome == TestOutcome.Succeeded) + " " + TestOutcome.Succeeded;
            var failed = _model.LastRun.Count(test => test.Result.Outcome == TestOutcome.Failed) + " " + TestOutcome.Failed;
            var inconclusive = _model.LastRun.Count(test => test.Result.Outcome == TestOutcome.Inconclusive) + " " + TestOutcome.Inconclusive;
            var resource = "Rubberduck Unit Tests - {0}\n{1} | {2} | {3}\n";
            var text = string.Format(resource, DateTime.Now, passed, failed, inconclusive) + results;

            _clipboard.Write(text);
        }
    }
}
