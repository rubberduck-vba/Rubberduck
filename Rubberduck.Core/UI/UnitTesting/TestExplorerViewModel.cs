using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.UI.Settings;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.UnitTesting;
using DataFormats = System.Windows.DataFormats;

namespace Rubberduck.UI.UnitTesting
{
    internal enum TestExplorerGrouping
    {
        None,
        Outcome,
        Category,
        Location
    }

    internal sealed class TestExplorerViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly ITestEngine _testEngine;
        private readonly IClipboardWriter _clipboard;
        private readonly ISettingsFormFactory _settingsFormFactory;

        public TestExplorerViewModel(RubberduckParserState state,
             ITestEngine testEngine,
             TestExplorerModel model,
             IClipboardWriter clipboard,
             IGeneralConfigService configService,
             ISettingsFormFactory settingsFormFactory)
        {
            _state = state;
            _testEngine = testEngine;
            _testEngine.TestCompleted += TestEngineTestCompleted;
            _clipboard = clipboard;
            _settingsFormFactory = settingsFormFactory;

            NavigateCommand = new NavigateCommand(_state.ProjectsProvider);  
            RunSingleTestCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteSingleTestCommand, CanExecuteSingleTestCommand);
            RunSelectedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteSelectedTestsCommand, CanExecuteSelectedTestsCommand);
            RunSelectedGroupCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunSelectedCategoryTestsCommand, CanExecuteRunSelectedCategoryTestsCommand);
            CopyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCopyResultsCommand);
            OpenTestSettingsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), OpenSettings);

            Model = model;
            if (CollectionViewSource.GetDefaultView(Model.Tests) is ListCollectionView tests)
            {
                tests.SortDescriptions.Add(new SortDescription("QualifiedName.QualifiedModuleName.Name", ListSortDirection.Ascending));
                tests.SortDescriptions.Add(new SortDescription("QualifiedName.MemberName", ListSortDirection.Ascending));
                Tests = tests;
            }

            OnPropertyChanged(nameof(Tests));
            TestGrouping = TestExplorerGrouping.Outcome;
        }

        public event EventHandler<TestCompletedEventArgs> TestCompleted;
        private void TestEngineTestCompleted(object sender, TestCompletedEventArgs e)
        {
            // Propagate the event
            TestCompleted?.Invoke(sender, e);
            Tests.Refresh();
        }

        public INavigateSource SelectedItem => SelectedTest;

        private TestMethodViewModel _selectedTest;
        public TestMethodViewModel SelectedTest
        {
            get => _selectedTest;
            set
            {
                _selectedTest = value;
                OnPropertyChanged();
            }
        }

        private static readonly Dictionary<TestExplorerGrouping, PropertyGroupDescription> GroupDescriptions = new Dictionary<TestExplorerGrouping, PropertyGroupDescription>
        {
            { TestExplorerGrouping.Outcome, new PropertyGroupDescription("Result.Outcome", new TestResultToOutcomeTextConverter()) },
            { TestExplorerGrouping.Location, new PropertyGroupDescription("QualifiedName.QualifiedModuleName.Name") },
            { TestExplorerGrouping.Category, new PropertyGroupDescription("Method.Category.Name") }
        };

        private TestExplorerGrouping _grouping = TestExplorerGrouping.None;

        public TestExplorerGrouping TestGrouping
        {
            get => _grouping;
            set
            {
                if (value == _grouping)
                {
                    return;
                }

                _grouping = value;
                Tests.GroupDescriptions.Clear();
                Tests.GroupDescriptions.Add(GroupDescriptions[_grouping]);
                Tests.Refresh();
                OnPropertyChanged();
            }
        }

        public ReparseCommand RefreshCommand { get; set; }

        public RunAllTestsCommand RunAllTestsCommand { get; set; }
        public RepeatLastRunCommand RepeatLastRunCommand { get; set; }
        public RunNotExecutedTestsCommand RunNotExecutedTestsCommand { get; set; }
        // no way to run skipped tests. Those are skipped until reparsing anyways, so it's k
        public RunInconclusiveTestsCommand RunInconclusiveTestsCommand { get; set; }
        public RunFailedTestsCommand RunFailedTestsCommand { get; set; }
        public RunSucceededTestsCommand RunPassedTestsCommand { get; set; }
        public CommandBase RunSingleTestCommand { get; }
        public CommandBase RunSelectedTestsCommand { get; }
        public CommandBase RunSelectedGroupCommand { get; }

        public AddTestModuleCommand AddTestModuleCommand { get; set; }
        public AddTestMethodCommand AddTestMethodCommand { get; set; }
        public AddTestMethodExpectedErrorCommand AddErrorTestMethodCommand { get; set; }

        public CommandBase CopyResultsCommand { get; }

        public CommandBase OpenTestSettingsCommand { get; }

        public INavigateCommand NavigateCommand { get; }

        private void OpenSettings(object param)
        {
            using (var window = _settingsFormFactory.Create(SettingsViews.UnitTestSettings))
            {
                window.ShowDialog();
                _settingsFormFactory.Release(window);
            }
        }

        public TestExplorerModel Model { get; }

        public ICollectionView Tests { get; }

        private bool CanExecuteSingleTestCommand(object obj)
        {
            return !Model.IsBusy && SelectedItem != null;
        }

        private bool CanExecuteSelectedTestsCommand(object obj)
        {
            return !Model.IsBusy && obj is IList viewModels && viewModels.Count > 0;
        }

        private void ExecuteSingleTestCommand(object obj)
        {
            if (SelectedTest == null)
            {
                return;
            }
            
            Model.IsBusy = true;
            _testEngine.Run(new[] { SelectedTest.Method });
            Model.IsBusy = false;
        }

        private void ExecuteSelectedTestsCommand(object obj)
        {
            if (Model.IsBusy || !(obj is IList viewModels && viewModels.Count > 0))
            {
                return;
            }

            var models = viewModels.OfType<TestMethodViewModel>().Select(model => model.Method).ToArray();

            if (!models.Any())
            {
                return;
            }

            Model.IsBusy = true;
            _testEngine.Run(models);
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

            using (var strm1 = ExportFormatter.XmlSpreadsheetNew(aResults, title, columnInfos))
            {
                //Add the formats from richest formatting to least formatting
                _clipboard.AppendStream(DataFormats.GetDataFormat(XML_SPREADSHEET_DATA_FORMAT).Name, strm1);
                _clipboard.AppendString(DataFormats.Rtf, rtfResults);
                _clipboard.AppendString(DataFormats.Html, htmlResults);
                _clipboard.AppendString(DataFormats.CommaSeparatedValue, csvResults);
                //_clipboard.AppendString(DataFormats.UnicodeText, textResults);

                _clipboard.Flush();
            }
        }

        private void ExecuteRunSelectedCategoryTestsCommand(object obj)
        {
            if (SelectedTest == null)
            {
                return;
            }
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
