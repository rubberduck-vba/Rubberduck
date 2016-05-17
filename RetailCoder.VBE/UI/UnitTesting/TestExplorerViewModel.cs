using System;
using System.Linq;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;
using Rubberduck.UI.Settings;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerViewModel : ViewModelBase, INavigateSelection
    {
        private readonly ITestEngine _testEngine;
        private readonly TestExplorerModel _model;
        private readonly IClipboardWriter _clipboard;
        private readonly IGeneralConfigService _configService;

        public TestExplorerViewModel(VBE vbe, RubberduckParserState state, ITestEngine testEngine, TestExplorerModel model, IClipboardWriter clipboard, NewUnitTestModuleCommand newTestModuleCommand, NewTestMethodCommand newTestMethodCommand, IGeneralConfigService configService)
        {
            _testEngine = testEngine;
            _testEngine.TestCompleted += TestEngineTestCompleted;
            _model = model;
            _clipboard = clipboard;
            _configService = configService;

            _navigateCommand = new NavigateCommand();

            _runAllTestsCommand = new RunAllTestsCommand(testEngine, model, state);
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

            _openTestSettingsCommand = new DelegateCommand(OpenSettings);
        }

        private bool CanExecuteRunPassedTestsCommand(object obj)
        {
            return _model.Tests.Any(test => test.Result.Outcome == TestOutcome.Succeeded);
        }

        private bool CanExecuteRunFailedTestsCommand(object obj)
        {
            return _model.Tests.Any(test => test.Result.Outcome == TestOutcome.Failed);
        }

        private bool CanExecuteRunNotExecutedTestsCommand(object obj)
        {
            return _model.Tests.Any(test => test.Result.Outcome == TestOutcome.Unknown);
        }

        private bool CanExecuteRepeatLastRunCommand(object obj)
        {
            return _model.LastRun.Any();
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

        public INavigateSource SelectedItem { get { return SelectedTest; } }

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

        private bool _groupByOutcome = true;
        public bool GroupByOutcome
        {
            get { return _groupByOutcome; }
            set
            {
                if (_groupByOutcome != value)
                {
                    _groupByOutcome = value;
                    OnPropertyChanged();
                }
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
        
        private readonly ICommand _openTestSettingsCommand;
        public ICommand OpenTestSettingsCommand { get { return _openTestSettingsCommand; } }

        private void OpenSettings(object param)
        {
            using (var window = new SettingsForm(_configService, SettingsViews.UnitTestSettings))
            {
                window.ShowDialog();
            }
        }

        public TestExplorerModel Model { get { return _model; } }

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
            return !Model.IsBusy && SelectedItem != null;
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
            var ignored = _model.LastRun.Count(test => test.Result.Outcome == TestOutcome.Ignored) + " " + TestOutcome.Ignored;
            var resource = "Rubberduck Unit Tests - {0}\n{1} | {2} | {3}\n";
            var text = string.Format(resource, DateTime.Now, passed, failed, inconclusive, ignored) + results;

            _clipboard.Write(text);
        }
    }
}
