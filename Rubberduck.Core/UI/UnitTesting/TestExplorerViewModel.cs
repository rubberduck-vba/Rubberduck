using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.UI.Settings;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ITestEngine _testEngine;
        private readonly IClipboardWriter _clipboard;
        private readonly ISettingsFormFactory _settingsFormFactory;
        private readonly IMessageBox _messageBox;

        public TestExplorerViewModel(IVBE vbe,
             RubberduckParserState state,
             ITestEngine testEngine,
             TestExplorerModel model,
             IClipboardWriter clipboard,
             IGeneralConfigService configService,
             ISettingsFormFactory settingsFormFactory,
             IMessageBox messageBox,
             ReparseCommand reparseCommand)
        {
            _vbe = vbe;
            _state = state;
            _testEngine = testEngine;
            _testEngine.TestCompleted += TestEngineTestCompleted;
            Model = model;
            _clipboard = clipboard;
            _settingsFormFactory = settingsFormFactory;
            _messageBox = messageBox;

            _navigateCommand = new NavigateCommand(_state.ProjectsProvider);

            RepeatLastRunCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRepeatLastRunCommand, CanExecuteRepeatLastRunCommand);
            RunNotExecutedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunNotExecutedTestsCommand, CanExecuteRunNotExecutedTestsCommand);
            RunInconclusiveTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunInconclusiveTestsCommand, CanExecuteRunInconclusiveTestsCommand);
            RunFailedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunFailedTestsCommand, CanExecuteRunFailedTestsCommand);
            RunPassedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunPassedTestsCommand, CanExecuteRunPassedTestsCommand);
            RunSelectedTestCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteSelectedTestCommand, CanExecuteSelectedTestCommand);
            RunSelectedCategoryTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunSelectedCategoryTestsCommand, CanExecuteRunSelectedCategoryTestsCommand);

            CopyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCopyResultsCommand);

            OpenTestSettingsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), OpenSettings);

            SetOutcomeGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByOutcome = true;

                if ((bool)param)
                {
                    GroupByLocation = false;
                    GroupByCategory = false;
                }
            });

            SetLocationGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByLocation = true;

                if ((bool)param)
                {
                    GroupByOutcome = false;
                    GroupByCategory = false;
                }
            });

            SetCategoryGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByCategory = true;

                if ((bool)param)
                {
                    GroupByOutcome = false;
                    GroupByLocation = false;
                }
            });
        }

        private static readonly ParserState[] AllowedRunStates = { ParserState.ResolvedDeclarations, ParserState.ResolvingReferences, ParserState.Ready };

        private bool CanExecuteRunPassedTestsCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) && Model.Tests.Any(test => test.Result.Outcome == TestOutcome.Succeeded);
        }

        private bool CanExecuteRunFailedTestsCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) && Model.Tests.Any(test => test.Result.Outcome == TestOutcome.Failed);
        }

        private bool CanExecuteRunNotExecutedTestsCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) && Model.Tests.Any(test => test.Result.Outcome == TestOutcome.Unknown);
        }

        private bool CanExecuteRunInconclusiveTestsCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) & Model.Tests.Any(test => test.Result.Outcome == TestOutcome.Inconclusive);
        }

        private bool CanExecuteRepeatLastRunCommand(object obj)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status) && Model.LastRun.Any();
        }

        public event EventHandler<EventArgs> TestCompleted;
        private void TestEngineTestCompleted(object sender, EventArgs e)
        {
            if (e is TestCompletedEventArgs args)
            {
                Model.Tests.First(vm => vm.Method == args.Test).Result = args.Result;
                // get the UI to update
                TestCompleted?.Invoke(sender, e);
            }
        }

        public INavigateSource SelectedItem => SelectedTest;

        private TestMethodViewModel _selectedTest;
        internal TestMethodViewModel SelectedTest
        {
            get => _selectedTest;
            set
            {
                _selectedTest = value;
                OnPropertyChanged();
            }
        }

        private bool _groupByOutcome = true;
        public bool GroupByOutcome
        {
            get => _groupByOutcome;
            set
            {
                if (_groupByOutcome == value)
                {
                    return;
                }

                _groupByOutcome = value;
                OnPropertyChanged();
            }
        }

        private bool _groupByLocation;
        public bool GroupByLocation
        {
            get => _groupByLocation;
            set
            {
                if (_groupByLocation == value)
                {
                    return;
                }

                _groupByLocation = value;
                OnPropertyChanged();
            }
        }

        private bool _groupByCategory;
        public bool GroupByCategory
        {
            get => _groupByCategory;
            set
            {
                if (_groupByCategory == value)
                {
                    return;
                }

                _groupByCategory = value;
                OnPropertyChanged();
            }
        }

        public CommandBase SetOutcomeGroupingCommand { get; }

        public CommandBase SetLocationGroupingCommand { get; }

        public CommandBase SetCategoryGroupingCommand { get; }
        
        public RunAllTestsCommand RunAllTestsCommand { get; set; }

        public AddTestModuleCommand AddTestModuleCommand { get; set; }

        public AddTestMethodCommand AddTestMethodCommand { get; set; }

        public AddTestMethodExpectedErrorCommand AddErrorTestMethodCommand { get; set; }

        public ReparseCommand RefreshCommand { get; set; }

        public CommandBase RepeatLastRunCommand { get; }

        public CommandBase RunNotExecutedTestsCommand { get; }

        public CommandBase RunInconclusiveTestsCommand { get; }

        public CommandBase RunFailedTestsCommand { get; }

        public CommandBase RunPassedTestsCommand { get; }

        public CommandBase CopyResultsCommand { get; }

        private readonly NavigateCommand _navigateCommand;
        public INavigateCommand NavigateCommand => _navigateCommand;

        public CommandBase RunSelectedTestCommand { get; }

        public CommandBase RunSelectedCategoryTestsCommand { get; }

        public CommandBase OpenTestSettingsCommand { get; }

        private void OpenSettings(object param)
        {
            using (var window = _settingsFormFactory.Create())
            {
                window.ShowDialog();
                _settingsFormFactory.Release(window);
            }
        }

        public TestExplorerModel Model { get; }

        private void ExecuteRepeatLastRunCommand(object parameter)
        {
            var tests = Model.LastRun.ToList();
            Model.ClearLastRun();
            Model.IsBusy = true;
            _testEngine.Run(tests.Select(vm => vm.Method));
            Model.IsBusy = false;
        }

        private void ExecuteRunNotExecutedTestsCommand(object parameter)
        {
            Model.ClearLastRun();
            Model.IsBusy = true;
            _testEngine.Run(Model.Tests.Where(test => test.Result.Outcome == TestOutcome.Unknown).Select(vm => vm.Method));
            Model.IsBusy = false;
        }

        private void ExecuteRunInconclusiveTestsCommand(object parameter)
        {
            Model.ClearLastRun();
            Model.IsBusy = true;
            _testEngine.Run(Model.Tests.Where(test => test.Result.Outcome == TestOutcome.Inconclusive).Select(vm => vm.Method));
            Model.IsBusy = false;
        }

        private void ExecuteRunFailedTestsCommand(object parameter)
        {
            Model.ClearLastRun();
            Model.IsBusy = true;
            _testEngine.Run(Model.Tests.Where(test => test.Result.Outcome == TestOutcome.Failed).Select(vm => vm.Method));
            Model.IsBusy = false;
        }

        private void ExecuteRunPassedTestsCommand(object parameter)
        {
            Model.ClearLastRun();
            Model.IsBusy = true;
            _testEngine.Run(Model.Tests.Where(test => test.Result.Outcome == TestOutcome.Succeeded).Select(vm => vm.Method));
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

            Model.ClearLastRun();
            Model.IsBusy = true;
            _testEngine.Run(new[] { SelectedTest.Method });
            Model.IsBusy = false;
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            const string XML_SPREADSHEET_DATA_FORMAT = "XML Spreadsheet";

            ColumnInfo[] columnInfos = { new ColumnInfo("Project"), new ColumnInfo("Component"), new ColumnInfo("Method"), new ColumnInfo("Outcome"), new ColumnInfo("Output"),
                                           new ColumnInfo("Start Time"), new ColumnInfo("End Time"), new ColumnInfo("Duration (ms)", hAlignment.Right) };

            // FIXME do that to the TestMethodViewModel
            var aResults = Model.Tests.Select(test => test.ToArray()).ToArray();

            var title = string.Format($"Rubberduck Test Results - {DateTime.Now.ToString(CultureInfo.InvariantCulture)}");

            //var textResults = title + Environment.NewLine + string.Join("", _results.Select(result => result.ToString() + Environment.NewLine).ToArray());
            var csvResults = ExportFormatter.Csv(aResults, title, columnInfos);
            var htmlResults = ExportFormatter.HtmlClipboardFragment(aResults, title, columnInfos);
            var rtfResults = ExportFormatter.RTF(aResults, title);

            var strm1 = ExportFormatter.XmlSpreadsheetNew(aResults, title, columnInfos);
            //Add the formats from richest formatting to least formatting
            _clipboard.AppendStream(DataFormats.GetDataFormat(XML_SPREADSHEET_DATA_FORMAT).Name, strm1);
            _clipboard.AppendString(DataFormats.Rtf, rtfResults);
            _clipboard.AppendString(DataFormats.Html, htmlResults);
            _clipboard.AppendString(DataFormats.CommaSeparatedValue, csvResults);
            //_clipboard.AppendString(DataFormats.UnicodeText, textResults);

            _clipboard.Flush();
        }

        private void ExecuteRunSelectedCategoryTestsCommand(object obj)
        {
            if (SelectedTest == null)
            {
                return;
            }
            Model.ClearLastRun();
            Model.IsBusy = true;
            _testEngine.Run(Model.Tests.Where(test => test.Method.Category.Equals(SelectedTest.Method.Category))
                .Select(t => t.Method));
            Model.IsBusy = false;
        }

        private bool CanExecuteRunSelectedCategoryTestsCommand(object obj)
        {
            if (Model.IsBusy || SelectedItem == null)
            {
                return false;
            }

            return ((TestMethod) SelectedItem).Category.Name != string.Empty;
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
            Model.Dispose();
        }
    }
}
