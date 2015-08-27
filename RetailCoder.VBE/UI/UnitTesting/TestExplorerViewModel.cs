using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Reflection;
using Rubberduck.Reflection;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;
using resx = Rubberduck.UI.RubberduckUI;

namespace Rubberduck.UI.UnitTesting
{
    public abstract class TestExplorerModelBase : ViewModelBase, INotifyCollectionChanged
    {
        protected TestExplorerModelBase(IDictionary<TestMethod, TestResult> tests = null)
        {
            Tests = tests ?? new Dictionary<TestMethod, TestResult>();
        }

        public abstract void Refresh();
        
        protected readonly IDictionary<TestMethod, TestResult> Tests;

        /// <summary>
        /// Adds a <see cref="TestMethod"/> to the <see cref="Tests"/> dictionary, with an <see cref="TestResult.Unknown"/> result.
        /// </summary>
        /// <param name="test">The <see cref="TestMethod"/> to add.</param>
        protected void AddTest(TestMethod test)
        {
            Tests.Add(test, TestResult.Unknown());
            OnPropertyChanged("AllTests");
            OnPropertyChanged("TestCount");
            OnPropertyChanged("TestItems");
            OnCollectionChanged();
        }

        public void SetResult(TestMethod test, TestResult result)
        {
            Tests[test] = result;
            OnPropertyChanged("AllTests");
            OnPropertyChanged("TestItems");
            OnCollectionChanged();
        }

        public IReadOnlyDictionary<TestMethod, TestResult> AllTests
        {
            get { return new ReadOnlyDictionary<TestMethod, TestResult>(Tests); }
        }

        public ObservableCollection<TestResultGrouping> TestItems
        {
            get
            {
                return new ObservableCollection<TestResultGrouping>(Tests.GroupBy(t => t.Value.Outcome)
                    .Select(outcome => new TestResultGrouping(outcome.Key, outcome.ToDictionary(kvp => kvp.Key, kvp => kvp.Value))));
            }
        }

        private static readonly string[] ReservedTestAttributeNames =
        {
            "ModuleInitialize",
            "TestInitialize", 
            "TestCleanup",
            "ModuleCleanup"
        };

        public int TestCount { get { return Tests.Count; } }
        public int ExecutedCount { get { return Tests.Count(kvp => kvp.Value.Outcome != TestOutcome.Unknown); } }

        public string FailedCount
        {
            get
            {
                return string.Format(resx.TestExplorer_TestNumberFailed,
                    Tests.Values.Count(test => test.Outcome == TestOutcome.Failed));
            }
        }

        public string SuccessfulCount
        {
            get
            {
                return string.Format(resx.TestExplorer_TestNumberFailed,
                    Tests.Values.Count(test => test.Outcome == TestOutcome.Succeeded));
            }
        }

        public string InconclusiveCount
        {
            get
            {
                return string.Format(resx.TestExplorer_TestNumberFailed,
                    Tests.Values.Count(test => test.Outcome == TestOutcome.Inconclusive));
            }
        }

        /// <summary>
        /// A method that determines whether a <see cref="Member"/> is a test method or not.
        /// </summary>
        /// <param name="member">The <see cref="Member"/> to evaluate.</param>
        /// <returns>Returns <c>true</c> if specified <see cref="member"/> is a test method.</returns>
        protected static bool IsTestMethod(Member member)
        {
            var isIgnoredMethod = member.HasAttribute<TestInitializeAttribute>()
                                  || member.HasAttribute<TestCleanupAttribute>()
                                  || member.HasAttribute<ModuleInitializeAttribute>()
                                  || member.HasAttribute<ModuleCleanupAttribute>()
                                  || (ReservedTestAttributeNames.Any(attribute =>
                                      member.QualifiedMemberName.MemberName.StartsWith(attribute)));

            var result = !isIgnoredMethod &&
                (member.QualifiedMemberName.MemberName.StartsWith("Test") || member.HasAttribute<TestMethodAttribute>())
                 && member.Signature.Contains(member.QualifiedMemberName.MemberName + "()")
                 && member.MemberType == MemberType.Sub
                 && member.MemberVisibility == MemberVisibility.Public;

            return result;
        }

        public event NotifyCollectionChangedEventHandler CollectionChanged;

        private void OnCollectionChanged()
        {
            var handler = CollectionChanged;
            if (handler != null)
            {
                handler(this, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
            }
        }
    }

    public class TestResultGrouping
    {
        private readonly TestOutcome _outcome;
        private readonly IDictionary<TestMethod, TestResult> _tests;

        public TestResultGrouping(TestOutcome outcome, IDictionary<TestMethod, TestResult> tests)
        {
            _outcome = outcome;
            _tests = tests;

            _label = _outcome == TestOutcome.Unknown
                    ? string.Concat(_tests.Count, " ", RubberduckUI.TestExplorer_RunNotRunTests)
                    : string.Format(RubberduckUI.ResourceManager.GetString("TestExplorer_TestNumber" + _outcome), _tests.Count);

            _icon = new TestOutcomeImageSourceConverter().Convert(_outcome, null, null, null) as ImageSource;

            _items = tests.Select(kvp => new TestItem(kvp.Key, kvp.Value));
        }

        private readonly string _label;
        public string Label { get { return _label; } }

        private readonly ImageSource _icon;
        public ImageSource Icon { get { return _icon; } }

        private readonly IEnumerable<TestItem> _items;
        public IEnumerable<TestItem> Items { get { return _items; } }
    }
    
    public class TestItem
    {
        private readonly TestMethod _test;
        private readonly TestResult _result;

        public TestItem(TestMethod test, TestResult result)
        {
            _test = test;
            _result = result;
        }

        public string TestName { get { return _test.QualifiedMemberName.ToString(); } }
        public string Message { get { return _result.Output; } }
        public long Duration { get { return _result.Duration; } }
    }

    /// <summary>
    /// A TestExplorer model that discovers unit tests in standard modules (.bas) marked with a '@TestModule marker.
    /// </summary>
    public class StandardModuleTestExplorerModel : TestExplorerModelBase
    {
        private readonly VBE _vbe;

        public StandardModuleTestExplorerModel(VBE vbe)
        {
            _vbe = vbe;
        }

        public override void Refresh()
        {
            Tests.Clear();
            var tests = _vbe.VBProjects.Cast<VBProject>()
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                .Where(component => component.CodeModule.HasAttribute<TestModuleAttribute>())
                .Select(component => new { Component = component, Members = component.GetMembers(vbext_ProcKind.vbext_pk_Proc).Where(IsTestMethod) })
                .SelectMany(component => component.Members.Select(method =>
                    new TestMethod(method.QualifiedMemberName, _vbe.HostApplication())));

            foreach (var test in tests)
            {
                AddTest(test);
            }
        }
    }

    /// <summary>
    /// A TestExplorer model that discovers unit tests in a 'ThisOutlookSession' document/class module.
    /// </summary>
    /// <remarks>
    /// We can *discover* unit test methods all we want... we can't run them.
    /// </remarks>
    public class ThisOutlookSessionTestExplorerModel : TestExplorerModelBase
    {
        private readonly VBE _vbe;

        public ThisOutlookSessionTestExplorerModel(VBE vbe)
        {
            _vbe = vbe;
        }

        public override void Refresh()
        {
            Tests.Clear();
            var tests = _vbe.ActiveVBProject.VBComponents.Cast<VBComponent>()
                .SingleOrDefault(component => component.Type == vbext_ComponentType.vbext_ct_Document)
                .GetMembers(vbext_ProcKind.vbext_pk_Proc).Where(IsTestMethod)
                .Select(method => new TestMethod(method.QualifiedMemberName, _vbe.HostApplication()));

            foreach (var test in tests)
            {
                AddTest(test);
            }
        }
    }

    public class TestExplorerViewModel : ViewModelBase
    {
        private readonly ITestEngine _testEngine;
        private readonly TestExplorerModelBase _model;

        public TestExplorerViewModel(VBE vbe, ITestEngine testEngine, TestExplorerModelBase model)
        {
            _testEngine = testEngine;
            _model = model;

            _runAllTestsCommand = new RunAllTestsCommand(testEngine);
            _addTestModuleCommand = new AddTestModuleCommand(vbe);
            _addTestMethodCommand = new AddTestMethodCommand(vbe);
            _addErrorTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe);

            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand, CanExecuteRefreshCommand);
            _repeatLastRunCommand = new DelegateCommand(ExecuteRepeatLastRunCommand, CanExecuteRepeatLastRunCommand);
            _runNotExecutedTestsCommand = new DelegateCommand(ExecuteRunNotExecutedTestsCommand, CanExecuteRunNotExecutedTestsCommand);
            _runFailedTestsCommand = new DelegateCommand(ExecuteRunFailedTestsCommand, CanExecuteRunFailedTestsCommand);
            _runPassedTestsCommand = new DelegateCommand(ExecuteRunPassedTestsCommand, CanExecuteRunPassedTestsCommand);

            _copyResultsCommand = new DelegateCommand(ExecuteCopyResultsCommand, CanExecuteCopyResultsCommand);
            _exportResultsCommand = new DelegateCommand(ExecuteExportResultsCommand, CanExecuteExportResultsCommand);
        }

        private TestItem _selectedItem;
        public TestItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                OnPropertyChanged("SelectedItem");
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

        private bool _isBusy;
        public bool IsBusy 
        { 
            get { return _isBusy; }
            private set
            {
                _isBusy = value; 
                OnPropertyChanged(); 
                CommandManager.InvalidateRequerySuggested();
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

        private bool CanExecuteRefreshCommand(object parameter)
        {
            return !IsBusy;
        }

        private void ExecuteRepeatLastRunCommand(object parameter)
        {
            IsBusy = true;
            _testEngine.Run(_model.AllTests.Where(kvp => kvp.Value.Outcome != TestOutcome.Unknown).Select(kvp => kvp.Key));
            IsBusy = false;
        }

        private bool CanExecuteRepeatLastRunCommand(object parameter)
        {
            return !_isBusy && _model.AllTests.Any(kvp => kvp.Value.Outcome != TestOutcome.Unknown);
        }

        private void ExecuteRunNotExecutedTestsCommand(object parameter)
        {
            IsBusy = true;
            _testEngine.Run(_model.AllTests.Where(kvp => kvp.Value.Outcome == TestOutcome.Unknown).Select(kvp => kvp.Key));
            IsBusy = false;
        }

        private bool CanExecuteRunNotExecutedTestsCommand(object parameter)
        {
            return !IsBusy && _model.AllTests.Any(kvp => kvp.Value.Outcome == TestOutcome.Unknown);
        }

        private void ExecuteRunFailedTestsCommand(object parameter)
        {
            IsBusy = true;
            _testEngine.Run(_model.AllTests.Where(kvp => kvp.Value.Outcome == TestOutcome.Failed).Select(kvp => kvp.Key));
            IsBusy = false;
        }

        private bool CanExecuteRunFailedTestsCommand(object parameter)
        {
            return !IsBusy && _model.AllTests.Any(kvp => kvp.Value.Outcome == TestOutcome.Failed);
        }

        private void ExecuteRunPassedTestsCommand(object parameter)
        {
            IsBusy = true;
            _testEngine.Run(_model.AllTests.Where(kvp => kvp.Value.Outcome == TestOutcome.Succeeded).Select(kvp => kvp.Key));
            IsBusy = false;
        }

        private bool CanExecuteRunPassedTestsCommand(object parameter)
        {
            return !IsBusy && _model.AllTests.Any(kvp => kvp.Value.Outcome == TestOutcome.Succeeded);
        }

        private void ExecuteExportResultsCommand(object parameter)
        {
            throw new NotImplementedException();
        }

        private bool CanExecuteExportResultsCommand(object parameter)
        {
            return HasExportableResults();
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            throw new NotImplementedException();
        }

        private bool CanExecuteCopyResultsCommand(object parameter)
        {
            return HasExportableResults();
        }

        private bool HasExportableResults()
        {
            return !IsBusy && _model.AllTests.Any(kvp => kvp.Value.Outcome != TestOutcome.Unknown);
        }
    }
}
