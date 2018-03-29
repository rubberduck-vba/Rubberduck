﻿using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using NLog;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
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
             IMessageBox messageBox)
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

            RunAllTestsCommand = new RunAllTestsCommand(vbe, state, testEngine, model, null);
            RunAllTestsCommand.RunCompleted += RunCompleted;

            AddTestModuleCommand = new AddTestModuleCommand(vbe, state, configService, _messageBox);
            AddTestMethodCommand = new AddTestMethodCommand(vbe, state);
            AddErrorTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe, state);

            RefreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRefreshCommand, CanExecuteRefreshCommand);
            RepeatLastRunCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRepeatLastRunCommand, CanExecuteRepeatLastRunCommand);
            RunNotExecutedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunNotExecutedTestsCommand, CanExecuteRunNotExecutedTestsCommand);
            RunInconclusiveTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunInconclusiveTestsCommand, CanExecuteRunInconclusiveTestsCommand);
            RunFailedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunFailedTestsCommand, CanExecuteRunFailedTestsCommand);
            RunPassedTestsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRunPassedTestsCommand, CanExecuteRunPassedTestsCommand);
            RunSelectedTestCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteSelectedTestCommand, CanExecuteSelectedTestCommand);

            CopyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCopyResultsCommand);

            OpenTestSettingsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), OpenSettings);

            SetOutcomeGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByOutcome = (bool)param;
                GroupByLocation = !(bool)param;
            });

            SetLocationGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
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
            var handler = TestCompleted;
            handler?.Invoke(sender, e);
        }

        public INavigateSource SelectedItem => SelectedTest;

        private TestMethod _selectedTest;
        public TestMethod SelectedTest
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

        public CommandBase SetOutcomeGroupingCommand { get; }

        public CommandBase SetLocationGroupingCommand { get; }

        public long TotalDuration { get; private set; }

        public RunAllTestsCommand RunAllTestsCommand { get; }

        public CommandBase AddTestModuleCommand { get; }

        public CommandBase AddTestMethodCommand { get; }

        public CommandBase AddErrorTestMethodCommand { get; }

        public CommandBase RefreshCommand { get; }

        public CommandBase RepeatLastRunCommand { get; }

        public CommandBase RunNotExecutedTestsCommand { get; }

        public CommandBase RunInconclusiveTestsCommand { get; }

        public CommandBase RunFailedTestsCommand { get; }

        public CommandBase RunPassedTestsCommand { get; }

        public CommandBase CopyResultsCommand { get; }

        private readonly NavigateCommand _navigateCommand;
        public INavigateCommand NavigateCommand => _navigateCommand;

        public CommandBase RunSelectedTestCommand { get; }

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

        private void ExecuteRefreshCommand(object parameter)
        {
            if (Model.IsBusy)
            {
                return;
            }

            Model.Refresh();
            SelectedTest = null;
        }

        private bool CanExecuteRefreshCommand(object parameter)
        {
            return !Model.IsBusy && _state.IsDirty();
        }

        private void EnsureRubberduckIsReferencedForEarlyBoundTests()
        {
            var projectIdsOfMembersUsingAddInLibrary = _state.DeclarationFinder.AllUserDeclarations
                .Where(member => member.AsTypeName == "Rubberduck.PermissiveAssertClass" 
                                    || member.AsTypeName == "Rubberduck.AssertClass")
                .Select(member => member.ProjectId)
                .ToHashSet();
            var projectsUsingAddInLibrary = _state.DeclarationFinder
                .UserDeclarations(DeclarationType.Project)
                .Where(declaration => projectIdsOfMembersUsingAddInLibrary.Contains(declaration.ProjectId))
                .Select(declaration => declaration.Project);

            foreach (var project in projectsUsingAddInLibrary)
            {
                project?.EnsureReferenceToAddInLibrary();
            }
        }

        private void ExecuteRepeatLastRunCommand(object parameter)
        {
            EnsureRubberduckIsReferencedForEarlyBoundTests();

            var tests = Model.LastRun.ToList();
            Model.ClearLastRun();

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

            Model.ClearLastRun();

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(Model.Tests.Where(test => test.Result.Outcome == TestOutcome.Unknown));
            stopwatch.Stop();

            Model.IsBusy = false;
            TotalDuration = stopwatch.ElapsedMilliseconds;
        }

        private void ExecuteRunInconclusiveTestsCommand(object parameter)
        {
            EnsureRubberduckIsReferencedForEarlyBoundTests();

            Model.ClearLastRun();

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(Model.Tests.Where(test => test.Result.Outcome == TestOutcome.Inconclusive));
            stopwatch.Stop();

            Model.IsBusy = false;
            TotalDuration = stopwatch.ElapsedMilliseconds;
        }

        private void ExecuteRunFailedTestsCommand(object parameter)
        {
            EnsureRubberduckIsReferencedForEarlyBoundTests();

            Model.ClearLastRun();

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(Model.Tests.Where(test => test.Result.Outcome == TestOutcome.Failed));
            stopwatch.Stop();

            Model.IsBusy = false;
            TotalDuration = stopwatch.ElapsedMilliseconds;
        }

        private void ExecuteRunPassedTestsCommand(object parameter)
        {
            EnsureRubberduckIsReferencedForEarlyBoundTests();

            Model.ClearLastRun();

            var stopwatch = new Stopwatch();
            Model.IsBusy = true;

            stopwatch.Start();
            _testEngine.Run(Model.Tests.Where(test => test.Result.Outcome == TestOutcome.Succeeded));
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

            Model.ClearLastRun();

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

            ColumnInfo[] columnInfos = { new ColumnInfo("Project"), new ColumnInfo("Component"), new ColumnInfo("Method"), new ColumnInfo("Outcome"), new ColumnInfo("Output"),
                                           new ColumnInfo("Start Time"), new ColumnInfo("End Time"), new ColumnInfo("Duration (ms)", hAlignment.Right) };

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
            RunAllTestsCommand.RunCompleted -= RunCompleted;
        }
    }
}
