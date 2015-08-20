using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Reflection;
using Rubberduck.Reflection;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.UnitTesting
{
    public interface ITestExplorerModel
    {
        /// <summary>
        /// A method that discovers unit tests.
        /// </summary>
        void Refresh();

        /// <summary>
        /// A method that assigns a result to a test.
        /// </summary>
        /// <param name="test">The test that was executed.</param>
        /// <param name="result">The test result.</param>
        void SetResult(TestMethod test, TestResult result);

        /// <summary>
        /// The unit tests discovered by the <see cref="Refresh"/> method, and their results (if any).
        /// </summary>
        IReadOnlyDictionary<TestMethod, TestResult> AllTests { get; } 
    }

    public abstract class TestExplorerModelBase : ViewModelBase, ITestExplorerModel
    {
        public abstract void Refresh();
        
        protected readonly IDictionary<TestMethod, TestResult> Tests = new Dictionary<TestMethod, TestResult>();

        /// <summary>
        /// Adds a <see cref="TestMethod"/> to the <see cref="Tests"/> dictionary, with an <see cref="TestResult.Unknown"/> result.
        /// </summary>
        /// <param name="test">The <see cref="TestMethod"/> to add.</param>
        protected void AddTest(TestMethod test)
        {
            Tests.Add(test, TestResult.Unknown());
            OnPropertyChanged("AllTests");
        }

        public void SetResult(TestMethod test, TestResult result)
        {
            Tests[test] = result;
            OnPropertyChanged("AllTests");
        }

        public IReadOnlyDictionary<TestMethod, TestResult> AllTests
        {
            get { return new ReadOnlyDictionary<TestMethod, TestResult>(Tests); }
        }

        private static readonly string[] ReservedTestAttributeNames =
        {
            "ModuleInitialize",
            "TestInitialize", 
            "TestCleanup",
            "ModuleCleanup"
        };

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
        private readonly VBE _vbe;
        private readonly ITestEngine _testEngine;
        private readonly ITestExplorerModel _model;

        public TestExplorerViewModel(VBE vbe, ITestEngine testEngine, ITestExplorerModel model)
        {
            _vbe = vbe;
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
            _exportResultsCommand = new DelegateCommand(ExecuteExportResultsCommand, CanExecuteExportResultsCommand);
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

        private readonly ICommand _exportResultsCommand;
        public ICommand ExportResultsCommand { get { return _exportResultsCommand; } }

        private bool _isBusy;
        public bool IsBusy 
        { 
            get { return _isBusy; }
            private set
            {
                _isBusy = value; 
                OnPropertyChanged("IsBusy"); 
                CommandManager.InvalidateRequerySuggested();
            } 
        }

        private void ExecuteRefreshCommand(object parameter)
        {
            if (!_isBusy)
            {
                return;
            }

            IsBusy = false;
            _model.Refresh();
            IsBusy = true;
        }

        private bool CanExecuteRefreshCommand(object parameter)
        {
            return IsBusy;
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
            return !IsBusy && _model.AllTests.Any(kvp => kvp.Value.Outcome != TestOutcome.Unknown);
        }
    }
}
