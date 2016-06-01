using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
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
    public class TestExplorerViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly ITestEngine _testEngine;
        private readonly TestExplorerModel _model;
        private readonly IClipboardWriter _clipboard;
        private readonly IGeneralConfigService _configService;
        private readonly IOperatingSystem _operatingSystem;

        public TestExplorerViewModel(VBE vbe,
             RubberduckParserState state,
             ITestEngine testEngine,
             TestExplorerModel model,
             IClipboardWriter clipboard,
             NewUnitTestModuleCommand newTestModuleCommand,
             NewTestMethodCommand newTestMethodCommand,
             IGeneralConfigService configService,
             IOperatingSystem operatingSystem)
        {
            _testEngine = testEngine;
            _testEngine.TestCompleted += TestEngineTestCompleted;
            _model = model;
            _clipboard = clipboard;
            _configService = configService;
            _operatingSystem = operatingSystem;

            _navigateCommand = new NavigateCommand();

            _runAllTestsCommand = new RunAllTestsCommand(state, testEngine, model);
            _runAllTestsCommand.RunCompleted += RunCompleted;

            _addTestModuleCommand = new AddTestModuleCommand(vbe, state, newTestModuleCommand);
            _addTestMethodCommand = new AddTestMethodCommand(vbe, state, newTestMethodCommand);
            _addErrorTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe, state, newTestMethodCommand);

            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand, CanExecuteRefreshCommand);
            _repeatLastRunCommand = new DelegateCommand(ExecuteRepeatLastRunCommand, CanExecuteRepeatLastRunCommand);
            _runNotExecutedTestsCommand = new DelegateCommand(ExecuteRunNotExecutedTestsCommand, CanExecuteRunNotExecutedTestsCommand);
            _runFailedTestsCommand = new DelegateCommand(ExecuteRunFailedTestsCommand, CanExecuteRunFailedTestsCommand);
            _runPassedTestsCommand = new DelegateCommand(ExecuteRunPassedTestsCommand, CanExecuteRunPassedTestsCommand);
            _runSelectedTestCommand = new DelegateCommand(ExecuteSelectedTestCommand, CanExecuteSelectedTestCommand);

            _copyResultsCommand = new DelegateCommand(ExecuteCopyResultsCommand);

            _openTestSettingsCommand = new DelegateCommand(OpenSettings);

            _setOutcomeGroupingCommand = new DelegateCommand(param =>
            {
                GroupByOutcome = (bool)param;
                GroupByLocation = !(bool)param;
            });

            _setLocationGroupingCommand = new DelegateCommand(param =>
            {
                GroupByLocation = (bool)param;
                GroupByOutcome = !(bool)param;
            });
        }

        private void RunCompleted(object sender, TestRunEventArgs e)
        {
            TotalDuration = e.Duration;
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

        private bool _groupByLocation;
        public bool GroupByLocation
        {
            get { return _groupByLocation; }
            set
            {
                if (_groupByLocation != value)
                {
                    _groupByLocation = value;
                    OnPropertyChanged();
                }
            }
        }

        private readonly ICommand _setOutcomeGroupingCommand;
        public ICommand SetOutcomeGroupingCommand { get { return _setOutcomeGroupingCommand; } }

        private readonly ICommand _setLocationGroupingCommand;
        public ICommand SetLocationGroupingCommand { get { return _setLocationGroupingCommand; } }

        public long TotalDuration { get; private set; }

        private readonly RunAllTestsCommand _runAllTestsCommand;
        public RunAllTestsCommand RunAllTestsCommand { get { return _runAllTestsCommand; } }

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
            using (var window = new SettingsForm(_configService, _operatingSystem, SettingsViews.UnitTestSettings))
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

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(tests);
            stopwatch.Stop();

            Model.IsBusy = false;
            TotalDuration = stopwatch.ElapsedMilliseconds;
        }

        private void ExecuteRunNotExecutedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(_model.Tests.Where(test => test.Result.Outcome == TestOutcome.Unknown));
            stopwatch.Stop();

            Model.IsBusy = false;
            TotalDuration = stopwatch.ElapsedMilliseconds;
        }

        private void ExecuteRunFailedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(_model.Tests.Where(test => test.Result.Outcome == TestOutcome.Failed));
            stopwatch.Stop();

            Model.IsBusy = false;
            TotalDuration = stopwatch.ElapsedMilliseconds;
        }

        private void ExecuteRunPassedTestsCommand(object parameter)
        {
            _model.ClearLastRun();

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(_model.Tests.Where(test => test.Result.Outcome == TestOutcome.Succeeded));
            stopwatch.Stop();

            Model.IsBusy = false;
            TotalDuration = stopwatch.ElapsedMilliseconds;
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

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(new[] { SelectedTest });
            stopwatch.Stop();

            Model.IsBusy = false;
            TotalDuration = stopwatch.ElapsedMilliseconds;
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            const string XML_SPREADSHEET_DATA_FORMAT = "XML Spreadsheet";

            ColumnInfo[] ColumnInfos = { new ColumnInfo("Project"), new ColumnInfo("Component"), new ColumnInfo("Method"), new ColumnInfo("Outcome"), new ColumnInfo("Output"),
                                           new ColumnInfo("Start Time"), new ColumnInfo("End Time"), new ColumnInfo("Duration (ms)", hAlignment.Right) };

            var aResults = _model.Tests.Select(test => test.ToArray()).ToArray();

            var resource = "Rubberduck Test Results - {0}";
            var title = string.Format(resource, DateTime.Now.ToString(CultureInfo.InvariantCulture));

            //var textResults = title + Environment.NewLine + string.Join("", _results.Select(result => result.ToString() + Environment.NewLine).ToArray());
            var csvResults = ExportFormatter.Csv(aResults, title, ColumnInfos);
            var htmlResults = ExportFormatter.HtmlClipboardFragment(aResults, title, ColumnInfos);
            var rtfResults = ExportFormatter.RTF(aResults, title);

            MemoryStream strm1 = ExportFormatter.XmlSpreadsheetNew(aResults, title, ColumnInfos);
            //Add the formats from richest formatting to least formatting
            _clipboard.AppendStream(DataFormats.GetDataFormat(XML_SPREADSHEET_DATA_FORMAT).Name, strm1);
            _clipboard.AppendString(DataFormats.Rtf, rtfResults);
            _clipboard.AppendString(DataFormats.Html, htmlResults);
            _clipboard.AppendString(DataFormats.CommaSeparatedValue, csvResults);
            //_clipboard.AppendString(DataFormats.UnicodeText, textResults);

            _clipboard.Flush();
        }

        //KEEP THIS, AS IT MAKES FOR THE BASIS OF A USEFUL *SUMMARY* REPORT
        //private void ExecuteCopyResultsCommand(object parameter)
        //{
        //    var results = string.Join(Environment.NewLine, _model.LastRun.Select(test => test.ToString()));

        //    var passed = _model.LastRun.Count(test => test.Result.Outcome == TestOutcome.Succeeded) + " " + TestOutcome.Succeeded;
        //    var failed = _model.LastRun.Count(test => test.Result.Outcome == TestOutcome.Failed) + " " + TestOutcome.Failed;
        //    var inconclusive = _model.LastRun.Count(test => test.Result.Outcome == TestOutcome.Inconclusive) + " " + TestOutcome.Inconclusive;
        //    var ignored = _model.LastRun.Count(test => test.Result.Outcome == TestOutcome.Ignored) + " " + TestOutcome.Ignored;

        //    var duration = RubberduckUI.UnitTest_TotalDuration + " - " + TotalDuration;

        //    var resource = "Rubberduck Unit Tests - {0}{6}{1} | {2} | {3} | {4}{6}{5} ms{6}";
        //    var text = string.Format(resource, DateTime.Now, passed, failed, inconclusive, ignored, duration, Environment.NewLine) + results;

        //    _clipboard.Write(text);
        //}

        public void Dispose()
        {
            _runAllTestsCommand.RunCompleted -= RunCompleted;
        }
    }
}
