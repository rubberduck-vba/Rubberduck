using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Settings;
using Rubberduck.UI.UnitTesting.ComCommands;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.UnitTesting
{
    public enum TestExplorerGrouping
    {
        None,
        Outcome,
        Category,
        Location
    }

    [Flags]
    public enum TestExplorerOutcomeFilter
    {
        None = 0,
        Unknown = 1,
        Fail = 1 << 1,
        Inconclusive = 1 << 2,
        Succeeded = 1 << 3,
        All = Unknown | Fail | Inconclusive | Succeeded
    }

    internal sealed class TestExplorerViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly IClipboardWriter _clipboard;
        private readonly ISettingsFormFactory _settingsFormFactory;

        public TestExplorerViewModel(ISelectionService selectionService,
            TestExplorerModel model,
            IClipboardWriter clipboard,
            // ReSharper disable once UnusedParameter.Local - left in place because it will likely be needed for app wide font settings, etc.
            IConfigurationService<Configuration> configService,
            ISettingsFormFactory settingsFormFactory,
            IRewritingManager rewritingManager,
            IAnnotationUpdater annotationUpdater)
        {
            _clipboard = clipboard;
            _settingsFormFactory = settingsFormFactory;

            NavigateCommand = new NavigateCommand(selectionService);  
            RunSingleTestCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteSingleTestCommand, CanExecuteSingleTest);
            RunSelectedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteSelectedTestsCommand, CanExecuteSelectedCommands);
            RunSelectedGroupCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunSelectedGroupCommand, CanExecuteGroupCommand);
            CancelTestRunCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCancelTestRunCommand, CanExecuteCancelTestRunCommand);
            ResetResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteResetResultsCommand, CanExecuteResetResultsCommand);
            CopyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCopyResultsCommand);
            OpenTestSettingsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), OpenSettings);
            CollapseAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCollapseAll);
            ExpandAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteExpandAll);
            IgnoreTestCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteIgnoreTestCommand, CanExecuteIgnoreTestCommand);
            UnignoreTestCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteUnignoreTestCommand, CanExecuteUnignoreTestCommand);
            IgnoreSelectedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteIgnoreSelectedTestsCommand, CanExecuteIgnoreSelectedTests);
            UnignoreSelectedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteUnignoreSelectedTestsCommand, CanExecuteUnignoreSelectedTests);
            IgnoreGroupCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteIgnoreGroupCommand, CanExecuteIgnoreGroupCommand);
            UnignoreGroupCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteUnignoreGroupCommand, CanExecuteUnignoreGroupCommand);

            RewritingManager = rewritingManager;
            AnnotationUpdater = annotationUpdater;

            Model = model;

            if (CollectionViewSource.GetDefaultView(Model.Tests) is ListCollectionView tests)
            {
                tests.SortDescriptions.Add(new SortDescription("QualifiedName.QualifiedModuleName.Name", ListSortDirection.Ascending));
                tests.SortDescriptions.Add(new SortDescription("QualifiedName.MemberName", ListSortDirection.Ascending));
                tests.IsLiveFiltering = true;
                tests.IsLiveGrouping = true;
                Tests = tests;
            }

            

            OnPropertyChanged(nameof(Tests));
            TestGrouping = TestExplorerGrouping.Outcome;

            OutcomeFilter = TestExplorerOutcomeFilter.All;
        }

        public TestExplorerModel Model { get; }

        public ICollectionView Tests { get; }

        public INavigateSource SelectedItem => MouseOverTest;

        private TestMethodViewModel _mouseOverTest;
        public TestMethodViewModel MouseOverTest
        {
            get => _mouseOverTest;
            set
            {
                if (ReferenceEquals(_mouseOverTest, value))
                {
                    return;
                }
                _mouseOverTest = value;
                OnPropertyChanged();
                RefreshContextMenu();
            }
        }

        private CollectionViewGroup _mouseOverGroup;
        public CollectionViewGroup MouseOverGroup
        {
            get => _mouseOverGroup;
            set
            {
                if (ReferenceEquals(_mouseOverGroup, value))
                {
                    return;
                }
                _mouseOverGroup = value;
                OnPropertyChanged();
                RefreshContextMenu();
            }
        }

        private void RefreshContextMenu()
        {
            OnPropertyChanged(nameof(CanExecuteUnignoreTestCommand));
            OnPropertyChanged(nameof(CanExecuteIgnoreTestCommand));
            OnPropertyChanged(nameof(CanExecuteUnignoreGroupCommand));
            OnPropertyChanged(nameof(CanExecuteIgnoreGroupCommand));
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

        private TestExplorerOutcomeFilter _outcomeFilter = TestExplorerOutcomeFilter.All;
        public TestExplorerOutcomeFilter OutcomeFilter
        {
            get => _outcomeFilter;
            set
            {
                if (value == _outcomeFilter)
                {
                    return;
                }

                _outcomeFilter = value;
                OnPropertyChanged();

                Tests.Filter = FilterResults;
            }
        }

        private string _testNameFilter = string.Empty;
        public string TestNameFilter
        {
            get => _testNameFilter;
            set
            {
                if (_testNameFilter != value)
                {
                    _testNameFilter = value;
                    OnPropertyChanged();
                    Tests.Filter = FilterResults;
                    OnPropertyChanged(nameof(Tests));
                }
            }
        }

        private bool _expanded;
        public bool ExpandedState
        {
            get => _expanded;
            set
            {
                _expanded = value;
                OnPropertyChanged();
            }
        }
        /// <summary>
        /// Filtering for displaying the correct tests.
        /// Uses both <see cref="OutcomeFilter"/> and <see cref="TestNameFilter"/>
        /// </summary>
        private bool FilterResults(object unitTest)
        {
            var testMethodViewModel = unitTest as TestMethodViewModel;

            var passesNameFilter = testMethodViewModel.QualifiedName.MemberName.ToUpper().Contains(TestNameFilter?.ToUpper() ?? string.Empty);

            Enum.TryParse(testMethodViewModel.Result.Outcome.ToString(), out TestExplorerOutcomeFilter convertedOutcome);
            var passesOutcomeFilter = (OutcomeFilter & convertedOutcome) == convertedOutcome;

            return passesNameFilter && passesOutcomeFilter;
        }

        public IRewritingManager RewritingManager { get; }
        public IAnnotationUpdater AnnotationUpdater { get; }

        private TestMethod _mousedOverTestMethod => ((TestMethodViewModel)SelectedItem).Method;
        public bool CanExecuteUnignoreTestCommand(object obj) => SelectedItem != null && _mousedOverTestMethod.IsIgnored;
        public bool CanExecuteIgnoreTestCommand(object obj) => SelectedItem != null && !_mousedOverTestMethod.IsIgnored;

        public bool CanExecuteIgnoreSelectedTests(object obj)
        {
            if (!Model.IsBusy && obj is IList viewModels && viewModels.Count > 0)
            {
                return viewModels.Cast<TestMethodViewModel>().Count(test => test.Method.IsIgnored) != viewModels.Count;
            }

            return false;
        }

        public bool CanExecuteUnignoreSelectedTests(object obj)
        {
            if (!Model.IsBusy && obj is IList viewModels && viewModels.Count > 0)
            {
                return viewModels.Cast<TestMethodViewModel>().Any(test => test.Method.IsIgnored);
            }

            return false;
        }

        public bool CanExecuteIgnoreGroupCommand(object obj)
        {
            var groupItems = MouseOverGroup?.Items
                             ?? GroupContainingSelectedTest(MouseOverTest).Items;

            return groupItems.Cast<TestMethodViewModel>().Count(test => test.Method.IsIgnored) != groupItems.Count;
        }

        public bool CanExecuteUnignoreGroupCommand(object obj)
        {
            var groupItems = MouseOverGroup?.Items
                             ?? GroupContainingSelectedTest(MouseOverTest)?.Items;
            
            return groupItems != null 
                   && groupItems.Cast<TestMethodViewModel>().Any(test => test.Method.IsIgnored);
        }
        
        #region Commands

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

        public CommandBase CancelTestRunCommand { get; }
        public CommandBase ResetResultsCommand { get; }

        public AddTestModuleCommand AddTestModuleCommand { get; set; }
        public AddTestMethodCommand AddTestMethodCommand { get; set; }
        public AddTestMethodExpectedErrorCommand AddErrorTestMethodCommand { get; set; }

        public CommandBase CopyResultsCommand { get; }

        public CommandBase OpenTestSettingsCommand { get; }

        public INavigateCommand NavigateCommand { get; }

        public CommandBase CollapseAllCommand { get; }
        public CommandBase ExpandAllCommand { get; }

        public CommandBase IgnoreTestCommand { get; }
        public CommandBase UnignoreTestCommand { get; }

        public CommandBase IgnoreGroupCommand { get; }
        public CommandBase UnignoreGroupCommand { get; }

        public CommandBase IgnoreSelectedTestsCommand { get; }
        public CommandBase UnignoreSelectedTestsCommand { get; }

        #endregion

        #region Delegates

        private bool CanExecuteSingleTest(object obj)
        {
            return !Model.IsBusy && MouseOverTest != null;
        }

        private bool CanExecuteSelectedCommands(object obj)
        {
            return !Model.IsBusy && obj is IList viewModels && viewModels.Count > 0;
        }

        private bool CanExecuteGroupCommand(object obj)
        {
            return !Model.IsBusy && (MouseOverTest != null || MouseOverGroup != null);
        }

        private bool CanExecuteResetResultsCommand(object obj)
        {
            return !Model.IsBusy && Tests.OfType<TestMethodViewModel>().Any(test => test.Result.Outcome != TestOutcome.Unknown);
        }

        private bool CanExecuteCancelTestRunCommand(object obj)
        {
            return Model.IsBusy;
        }

        private void ExecuteCollapseAll(object parameter)
        {
            ExpandedState = false;
        }

        private void ExecuteExpandAll(object parameter)
        {
            ExpandedState = true;
        }

        private void ExecuteSingleTestCommand(object obj)
        {
            if (MouseOverTest == null)
            {
                return;
            }

            Model.ExecuteTests(new List<TestMethodViewModel> { MouseOverTest });
        }

        private void ExecuteSelectedTestsCommand(object obj)
        {
            if (Model.IsBusy || !(obj is IList viewModels && viewModels.Count > 0))
            {
                return;
            }

            var models = viewModels.OfType<TestMethodViewModel>().ToList();

            if (!models.Any())
            {
                return;
            }

            Model.ExecuteTests(models);
        }

        private void ExecuteRunSelectedGroupCommand(object obj)
        {
            var tests = GroupContainingSelectedTest(MouseOverTest);

            if (tests is null)
            {
                return;
            }

            Model.ExecuteTests(tests.Items.OfType<TestMethodViewModel>().ToList());
        }

        private CollectionViewGroup GroupContainingSelectedTest(TestMethodViewModel selectedTest)
        {
            return selectedTest is null
                ? MouseOverGroup
                : Tests.Groups.OfType<CollectionViewGroup>().FirstOrDefault(group => group.Items.Contains(selectedTest));
        }

        private void ExecuteCancelTestRunCommand(object parameter)
        {
            Model.CancelTestRun();
        }

        private void ExecuteResetResultsCommand(object parameter)
        {
            foreach (var test in Tests.OfType<TestMethodViewModel>())
            {
                test.Result = new TestResult(TestOutcome.Unknown);
            }

            Tests.Refresh();
        }

        private void ExecuteIgnoreTestCommand(object parameter)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();

            var testMethod = parameter is null
                ? _mousedOverTestMethod
                : (TestMethod)parameter;

            AnnotationUpdater.AddAnnotation(rewriteSession, testMethod.Declaration, new IgnoreTestAnnotation());

            rewriteSession.TryRewrite();
        }

        private void ExecuteUnignoreTestCommand(object parameter)
        {
            var testMethod = parameter is null
                ? _mousedOverTestMethod
                : (TestMethod)parameter;

            var ignoreTestAnnotations = testMethod.Declaration.Annotations
                .Where(pta => pta.Annotation is IgnoreTestAnnotation);

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();

            AnnotationUpdater.RemoveAnnotations(rewriteSession, ignoreTestAnnotations);

            rewriteSession.TryRewrite();
        }

        private void ExecuteIgnoreGroupCommand(object parameter)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var testGroup = GroupContainingSelectedTest(MouseOverTest);
            var ignoreTestAnnotation = new IgnoreTestAnnotation();
            
            foreach (TestMethodViewModel test in testGroup.Items)
            {
                if (!test.Method.IsIgnored)
                {
                    AnnotationUpdater.AddAnnotation(rewriteSession, test.Method.Declaration, ignoreTestAnnotation);
                }
            }

            rewriteSession.TryRewrite();
        }

        private void ExecuteUnignoreGroupCommand(object parameter)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var testGroup = GroupContainingSelectedTest(MouseOverTest);

            foreach (TestMethodViewModel test in testGroup.Items)
            {
                var ignoreTestAnnotations = test.Method.Declaration.Annotations
                    .Where(pta => pta.Annotation is IgnoreTestAnnotation);

                foreach (var ignoreTestAnnotation in ignoreTestAnnotations)
                {
                    AnnotationUpdater.RemoveAnnotation(rewriteSession, ignoreTestAnnotation);
                }
            }

            rewriteSession.TryRewrite();
        }

        private void ExecuteUnignoreSelectedTestsCommand(object parameter)
        {
            if (Model.IsBusy || !(parameter is IList viewModels && viewModels.Count > 0))
            {
                return;
            }

            var ignoredModels = viewModels.OfType<TestMethodViewModel>()
                .Where(model => model.Method.IsIgnored);

            if (!ignoredModels.Any())
            {
                return; 
            }

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();

            foreach (var test in ignoredModels)
            {
                var ignoreTestAnnotations = test.Method.Declaration.Annotations
                    .Where(pta => pta.Annotation is IgnoreTestAnnotation);

                AnnotationUpdater.RemoveAnnotations(rewriteSession, ignoreTestAnnotations);
            }

            rewriteSession.TryRewrite();
        }

        private void ExecuteIgnoreSelectedTestsCommand(object parameter)
        {
            if (Model.IsBusy || !(parameter is IList viewModels && viewModels.Count > 0))
            {
                return;
            }

            var unignoredModels = viewModels.OfType<TestMethodViewModel>()
                .Where(model => !model.Method.IsIgnored);

            if (!unignoredModels.Any())
            {
                return;
            }

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var ignoreTestAnnotation = new IgnoreTestAnnotation();

            foreach (var test in unignoredModels)
            {
                AnnotationUpdater.AddAnnotation(rewriteSession, test.Method.Declaration, ignoreTestAnnotation);
            }

            rewriteSession.TryRewrite();
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            const string XML_SPREADSHEET_DATA_FORMAT = "XML Spreadsheet";

            ColumnInfo[] columnInfos = { new ColumnInfo("Project"), new ColumnInfo("Component"), new ColumnInfo("Method"), new ColumnInfo("Outcome"), new ColumnInfo("Output"),
                                           new ColumnInfo("Start Time"), new ColumnInfo("End Time"), new ColumnInfo("Duration (ms)", hAlignment.Right) };

            // FIXME do that to the TestMethodViewModel
            var aResults = Model.Tests.Select(test => test.ToArray()).ToArray();

            var title = string.Format($"Rubberduck Test Results - {DateTime.Now.ToString(CultureInfo.InvariantCulture)}");

            var textResults = title + Environment.NewLine + string.Join(string.Empty, aResults.Select(result => result.ToString() + Environment.NewLine).ToArray());
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
                _clipboard.AppendString(DataFormats.UnicodeText, textResults);

                _clipboard.Flush();
            }
        }

        // TODO - FIXME
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

        private void OpenSettings(object param)
        {
            using (var window = _settingsFormFactory.Create(SettingsViews.UnitTestSettings))
            {
                window.ShowDialog();
                _settingsFormFactory.Release(window);
            }
        }

        #endregion

        public void Dispose()
        {
            Model.Dispose();
        }
    }
}
