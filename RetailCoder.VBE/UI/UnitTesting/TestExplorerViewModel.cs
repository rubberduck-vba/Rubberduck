using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using NLog;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;
using Rubberduck.UI.Settings;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ITestEngine _testEngine;
        private readonly TestExplorerModel _model;
        private readonly IClipboardWriter _clipboard;
        private readonly IGeneralConfigService _configService;
        private readonly IOperatingSystem _operatingSystem;

        public TestExplorerViewModel(IVBE vbe,
             RubberduckParserState state,
             ITestEngine testEngine,
             TestExplorerModel model,
             IClipboardWriter clipboard,
             IGeneralConfigService configService,
             IOperatingSystem operatingSystem)
        {
            _vbe = vbe;
            _state = state;
            _testEngine = testEngine;
            _testEngine.TestCompleted += TestEngineTestCompleted;
            _model = model;
            _clipboard = clipboard;
            _configService = configService;
            _operatingSystem = operatingSystem;

            _navigateCommand = new NavigateCommand();

            _runAllTestsCommand = new RunAllTestsCommand(vbe, state, testEngine, model, null);
            _runAllTestsCommand.RunCompleted += RunCompleted;

            _addTestModuleCommand = new AddTestModuleCommand(vbe, state, configService);
            _addTestMethodCommand = new AddTestMethodCommand(vbe, state);
            _addErrorTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe, state);

            _refreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRefreshCommand, CanExecuteRefreshCommand);
            _repeatLastRunCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRepeatLastRunCommand, CanExecuteRepeatLastRunCommand);
            _runNotExecutedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunNotExecutedTestsCommand, CanExecuteRunNotExecutedTestsCommand);
            _runFailedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunFailedTestsCommand, CanExecuteRunFailedTestsCommand);
            _runPassedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunPassedTestsCommand, CanExecuteRunPassedTestsCommand);
            _runSelectedTestCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteSelectedTestCommand, CanExecuteSelectedTestCommand);

            _copyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCopyResultsCommand);

            _openTestSettingsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), OpenSettings);

            _setOutcomeGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByOutcome = (bool)param;
                GroupByLocation = !(bool)param;
            });

            _setLocationGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByLocation = (bool)param;
                GroupByOutcome = !(bool)param;
            });
        }

        private void RunCompleted(object sender, TestRunEventArgs e)
        {
            TotalDuration = e.Duration;
        }

        private static readonly ParserState[] AllowedRunStates = { ParserState.ResolvedDeclarations, ParserState.ResolvingReferences, ParserState.Ready };

        private bool CanExecuteRunPassedTestsCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) && _model.Tests.Any(test => test.Result.Outcome == TestOutcome.Succeeded);
        }

        private bool CanExecuteRunFailedTestsCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) && _model.Tests.Any(test => test.Result.Outcome == TestOutcome.Failed);
        }

        private bool CanExecuteRunNotExecutedTestsCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) && _model.Tests.Any(test => test.Result.Outcome == TestOutcome.Unknown);
        }

        private bool CanExecuteRepeatLastRunCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) && _model.LastRun.Any();
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

        private readonly CommandBase _setOutcomeGroupingCommand;
        public CommandBase SetOutcomeGroupingCommand { get { return _setOutcomeGroupingCommand; } }

        private readonly CommandBase _setLocationGroupingCommand;
        public CommandBase SetLocationGroupingCommand { get { return _setLocationGroupingCommand; } }

        public long TotalDuration { get; private set; }

        private readonly RunAllTestsCommand _runAllTestsCommand;
        public RunAllTestsCommand RunAllTestsCommand { get { return _runAllTestsCommand; } }

        private readonly CommandBase _addTestModuleCommand;
        public CommandBase AddTestModuleCommand { get { return _addTestModuleCommand; } }

        private readonly CommandBase _addTestMethodCommand;
        public CommandBase AddTestMethodCommand { get { return _addTestMethodCommand; } }

        private readonly CommandBase _addErrorTestMethodCommand;
        public CommandBase AddErrorTestMethodCommand { get { return _addErrorTestMethodCommand; } }

        private readonly CommandBase _refreshCommand;
        public CommandBase RefreshCommand { get { return _refreshCommand; } }

        private readonly CommandBase _repeatLastRunCommand;
        public CommandBase RepeatLastRunCommand { get { return _repeatLastRunCommand; } }

        private readonly CommandBase _runNotExecutedTestsCommand;
        public CommandBase RunNotExecutedTestsCommand { get { return _runNotExecutedTestsCommand; } }

        private readonly CommandBase _runFailedTestsCommand;
        public CommandBase RunFailedTestsCommand { get { return _runFailedTestsCommand; } }

        private readonly CommandBase _runPassedTestsCommand;
        public CommandBase RunPassedTestsCommand { get { return _runPassedTestsCommand; } }

        private readonly CommandBase _copyResultsCommand;
        public CommandBase CopyResultsCommand { get { return _copyResultsCommand; } }

        private readonly NavigateCommand _navigateCommand;
        public INavigateCommand NavigateCommand { get { return _navigateCommand; } }

        private readonly CommandBase _runSelectedTestCommand;
        public CommandBase RunSelectedTestCommand { get { return _runSelectedTestCommand; } }

        private readonly CommandBase _openTestSettingsCommand;
        public CommandBase OpenTestSettingsCommand { get { return _openTestSettingsCommand; } }

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
            return !Model.IsBusy && _state.IsDirty();
        }

        private void EnsureRubberduckIsReferencedForEarlyBoundTests()
        {
            foreach (var member in _state.AllUserDeclarations)
            {
                if (member.AsTypeName == "Rubberduck.PermissiveAssertClass" ||
                    member.AsTypeName == "Rubberduck.AssertClass")
                {
                    member.Project.EnsureReferenceToAddInLibrary();
                }
            }
        }

        private void ExecuteRepeatLastRunCommand(object parameter)
        {
            EnsureRubberduckIsReferencedForEarlyBoundTests();

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
            EnsureRubberduckIsReferencedForEarlyBoundTests();

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
            EnsureRubberduckIsReferencedForEarlyBoundTests();

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
            EnsureRubberduckIsReferencedForEarlyBoundTests();

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

            EnsureRubberduckIsReferencedForEarlyBoundTests();

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
