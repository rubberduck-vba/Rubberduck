using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Vbe.Interop;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;
using resx = Rubberduck.UI.RubberduckUI;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerViewModel : ViewModelBase
    {
        private readonly ITestEngine _testEngine;
        private readonly TestExplorerModelBase _model;

        public TestExplorerViewModel(VBE vbe, ITestEngine testEngine, TestExplorerModelBase model)
        {
            _testEngine = testEngine;
            _testEngine.TestCompleted += TestEngineTestCompleted;
            _model = model;

            _navigateCommand = new NavigateCommand();

            _runAllTestsCommand = new RunAllTestsCommand(testEngine, model);
            _addTestModuleCommand = new AddTestModuleCommand(vbe);
            _addTestMethodCommand = new AddTestMethodCommand(vbe, model);
            _addErrorTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe, model);

            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand, CanExecuteRefreshCommand);
            _repeatLastRunCommand = new DelegateCommand(ExecuteRepeatLastRunCommand, CanExecuteRepeatLastRunCommand);
            _runNotExecutedTestsCommand = new DelegateCommand(ExecuteRunNotExecutedTestsCommand, CanExecuteRunNotExecutedTestsCommand);
            _runFailedTestsCommand = new DelegateCommand(ExecuteRunFailedTestsCommand, CanExecuteRunFailedTestsCommand);
            _runPassedTestsCommand = new DelegateCommand(ExecuteRunPassedTestsCommand, CanExecuteRunPassedTestsCommand);
            _runSelectedTestCommand = new DelegateCommand(ExecuteSelectedTestCommand, CanExecuteSelectedTestCommand);

            _copyResultsCommand = new DelegateCommand(ExecuteCopyResultsCommand);
            _exportResultsCommand = new DelegateCommand(ExecuteExportResultsCommand);
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

        private TestMethod _selectedItem;
        public TestMethod SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
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

        private readonly ICommand _exportResultsCommand;
        public ICommand ExportResultsCommand { get { return _exportResultsCommand; } }

        private readonly NavigateCommand _navigateCommand;
        public ICommand NavigateCommand { get { return _navigateCommand; } }

        private readonly ICommand _runSelectedTestCommand;
        public ICommand RunSelectedTestCommand { get { return _runSelectedTestCommand; } }

        private bool _isBusy;
        public bool IsBusy 
        { 
            get { return _isBusy; }
            private set
            {
                _isBusy = value; 
                OnPropertyChanged(); 
            } 
        }

        public TestExplorerModelBase Model { get { return _model; } }

        private void ExecuteRefreshCommand(object parameter)
        {
            if (_isBusy)
            {
                return;
            }

            IsBusy = true;
            _model.Refresh();
            SelectedItem = null;
            IsBusy = false;
        }

        private void EvaluateCanExecute()
        {
            Dispatcher.CurrentDispatcher.Invoke(CommandManager.InvalidateRequerySuggested);
        }

        private bool CanExecuteRefreshCommand(object parameter)
        {
            return !IsBusy;
        }

        private void ExecuteRepeatLastRunCommand(object parameter)
        {
            var tests = _model.LastRun.ToList();
            _model.ClearLastRun();

            IsBusy = true;
            _testEngine.Run(tests);
            IsBusy = false;
            EvaluateCanExecute();
        }

        private void ExecuteRunNotExecutedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            IsBusy = true;
            _testEngine.Run(_model.Tests.Where(test => test.Result.Outcome == TestOutcome.Unknown));
            IsBusy = false;
            EvaluateCanExecute();
        }

        private void ExecuteRunFailedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            IsBusy = true;
            _testEngine.Run(_model.Tests.Where(test => test.Result.Outcome == TestOutcome.Failed));
            IsBusy = false;
            EvaluateCanExecute();
        }

        private void ExecuteRunPassedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            IsBusy = true;
            _testEngine.Run(_model.Tests.Where(test => test.Result.Outcome == TestOutcome.Succeeded));
            IsBusy = false;
            EvaluateCanExecute();
        }

        private bool CanExecuteSelectedTestCommand(object obj)
        {
            return true; //SelectedItem != null;
        }

        private void ExecuteSelectedTestCommand(object obj)
        {
            if (SelectedItem == null)
            {
                return;
            }

            _model.ClearLastRun();

            IsBusy = true;
            _testEngine.Run(new[]{SelectedItem});
            IsBusy = false;
            EvaluateCanExecute();
        }

        private void ExecuteExportResultsCommand(object parameter)
        {
            throw new NotImplementedException();
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            throw new NotImplementedException();
        }
    }
}
