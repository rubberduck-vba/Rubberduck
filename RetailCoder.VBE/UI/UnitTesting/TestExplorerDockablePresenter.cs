using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Reflection;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.UnitTesting
{
    public interface ITestExplorerPresenter : IPresenter
    {
        void RunTests();
        void RunTests(IEnumerable<TestMethod> tests);
    }

    public class TestExplorerDockablePresenter : DockablePresenterBase, ITestExplorerPresenter
    {
        private readonly GridViewSort<TestExplorerItem> _gridViewSort;
        private readonly ITestEngine _testEngine;
        private readonly ITestExplorerWindow _view;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public TestExplorerDockablePresenter(VBE vbe, AddIn addin, ITestExplorerWindow control, ITestEngine testEngine, ICodePaneWrapperFactory wrapperFactory)
            : base(vbe, addin, control)
        {
            _testEngine = testEngine;
            _gridViewSort = new GridViewSort<TestExplorerItem>(RubberduckUI.Result, false);
            _wrapperFactory = wrapperFactory;

            _testEngine.ModuleInitialize += _testEngine_ModuleInitialize;
            _testEngine.ModuleCleanup += _testEngine_ModuleCleanup;
            _testEngine.MethodInitialize += TestEngineMethodInitialize;
            _testEngine.MethodCleanup += TestEngineMethodCleanup;

            _view = control; 
            _view.SortColumn += SortColumn;

            RegisterTestExplorerEvents();
        }

        private void SortColumn(object sender, DataGridViewCellMouseEventArgs e)
        {
            var columnName = _view.GridView.Columns[e.ColumnIndex].Name;

            // type "Image" doesn't implement "IComparable", so we need to sort by the outcome instead
            if (columnName == RubberduckUI.Result) { columnName = RubberduckUI.Outcome; }
            _view.AllTests = new BindingList<TestExplorerItem>(_gridViewSort.Sort(_view.AllTests.AsEnumerable(), columnName).ToList());
        }


        private void TestEngineMethodCleanup(object sender, TestModuleEventArgs e)
        {
            var module = e.QualifiedModuleName.Component.CodeModule;
            module.Parent.RunMethodsWithAttribute<TestCleanupAttribute>();
        }

        private void TestEngineMethodInitialize(object sender, TestModuleEventArgs e)
        {
            var module = e.QualifiedModuleName.Component.CodeModule;
            module.Parent.RunMethodsWithAttribute<TestInitializeAttribute>();
        }

        private void _testEngine_ModuleCleanup(object sender, TestModuleEventArgs e)
        {
            var module = e.QualifiedModuleName.Component.CodeModule;
            module.Parent.RunMethodsWithAttribute<ModuleCleanupAttribute>();
        }

        private void _testEngine_ModuleInitialize(object sender, TestModuleEventArgs e)
        {
            var module = e.QualifiedModuleName.Component.CodeModule;
            module.Parent.RunMethodsWithAttribute<ModuleInitializeAttribute>();
        }

        private void Synchronize()
        {
            FindAllTests();
            var results = new BindingList<TestExplorerItem>(_testEngine.Model.AllTests.Select(test => new TestExplorerItem(test.Key, test.Value)).ToList());
            _view.AllTests =
                new BindingList<TestExplorerItem>(
                    _gridViewSort.Sort(results, _gridViewSort.ColumnName,
                        _gridViewSort.SortedAscending).ToList());
        }

        public override void Show()
        {
            Synchronize();
            base.Show();
        }

        private void FindAllTests()
        {
            try
            {
                _testEngine.Model.Refresh();
            }
            catch (ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show(
                    RubberduckUI.TestExplorerDockablePresenter_MultipleTestsSameNameError,
                    RubberduckUI.Warning, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
        }

        public void RunTests()
        {
            RunTests(_testEngine.Model.AllTests.Keys);
        }

        public void RunTests(IEnumerable<TestMethod> tests)
        {
            try
            {
                _view.ClearResults();

                var testMethods = tests as IList<TestMethod> ?? tests.ToList(); //bypasses multiple enumeration
                _view.SetPlayList(testMethods);

                _view.ClearProgress();

                var projects = testMethods.Select(t => t.QualifiedMemberName.QualifiedModuleName.Project).Distinct();
                foreach (var project in projects)
                {
                    project.EnsureReferenceToAddInLibrary();
                }
            
                _testEngine.Run(testMethods);
            }
            catch (Exception exception)
            {
                // WTF is going on here?
            }
        }

        private void TestComplete(object sender, TestCompletedEventArgs e)
        {
            _view.WriteResult(e.Test, e.Result);
        }

        private void OnExplorerRefreshListButtonClick(object sender, EventArgs e)
        {
            Synchronize();
        }

        private void OnExplorerRunAllTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.Model.AllTests.Keys);
        }

        private void OnExplorerRunFailedTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.Model.AllTests.Where(test => test.Value.Outcome == TestOutcome.Failed).Select(kvp => kvp.Key));
        }

        private void OnExplorerRunLastRunTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.Model.AllTests.Where(test => test.Value.Outcome != TestOutcome.Unknown).Select(kvp => kvp.Key));
        }

        private void OnExplorerRunNotRunTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.Model.AllTests.Where(test => test.Value.Outcome == TestOutcome.Unknown).Select(kvp => kvp.Key));
        }

        private void OnExplorerRunPassedTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.Model.AllTests.Where(test => test.Value.Outcome == TestOutcome.Succeeded).Select(kvp => kvp.Key));
        }

        private void OnExplorerRunSelectedTestButtonClick(object sender, SelectedTestEventArgs e)
        {
            RunTests(e.Selection);
        }

        private void OnExplorerGoToSelectedTest(object sender, SelectedTestEventArgs e)
        {
            var controlSelection = e.Selection.FirstOrDefault();
            if (controlSelection == null)
            {
                return;
            }

            var startLine = 1;
            var startColumn = 1;
            var endLine = -1;
            var endColumn = -1;

            var signature = string.Concat("Public Sub ", controlSelection.QualifiedMemberName.MemberName, "()");

            var vbProject = VBE.VBProjects.Cast<VBProject>()
                    .FirstOrDefault(project => project.Protection != vbext_ProjectProtection.vbext_pp_locked
                                               && project.Equals(controlSelection.QualifiedMemberName.QualifiedModuleName.Project)
                                               && project.VBComponents
                                                   .Cast<VBComponent>()
                                                   .Any(c => c.Equals(controlSelection.QualifiedMemberName.QualifiedModuleName.Component)));

            if (vbProject == null)
            {
                return;
            }

            var vbComponent = vbProject.VBComponents.Cast<VBComponent>()
                                     .SingleOrDefault(component => component.Equals(controlSelection.QualifiedMemberName.QualifiedModuleName.Component));

            if (vbComponent == null)
            {
                return;
            }

            var codeModule = vbComponent.CodeModule;
            if (codeModule == null)
            {
                return;
            }

            if (codeModule.Find(signature, ref startLine, ref startColumn, ref endLine, ref endColumn))
            {
                var codePane = _wrapperFactory.Create(codeModule.CodePane);
                var selection = new Selection(startLine, startColumn, endLine, endColumn);
                codePane.Selection = selection;
            }
        }

        private void OnExplorerAddExpectedErrorTestMethodButtonClick(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewExpectedErrorTestMethod(this.VBE);
            Synchronize();
        }

        private void OnExplorerAddTestMethodButtonClick(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewTestMethod(this.VBE);
            Synchronize();
        }

        private void OnExplorerAddTestModuleButtonClick(object sender, EventArgs e)
        {
            NewUnitTestModuleCommand.NewUnitTestModule(this.VBE);
            Synchronize();
        }

        private void RegisterTestExplorerEvents()
        {
            _view.OnRefreshListButtonClick += OnExplorerRefreshListButtonClick;

            _view.OnRunAllTestsButtonClick += OnExplorerRunAllTestsButtonClick;
            _view.OnRunFailedTestsButtonClick += OnExplorerRunFailedTestsButtonClick;
            _view.OnRunLastRunTestsButtonClick += OnExplorerRunLastRunTestsButtonClick;
            _view.OnRunNotRunTestsButtonClick += OnExplorerRunNotRunTestsButtonClick;
            _view.OnRunPassedTestsButtonClick += OnExplorerRunPassedTestsButtonClick;
            _view.OnRunSelectedTestButtonClick += OnExplorerRunSelectedTestButtonClick;

            _view.OnGoToSelectedTest += OnExplorerGoToSelectedTest;

            _view.OnAddExpectedErrorTestMethodButtonClick += OnExplorerAddExpectedErrorTestMethodButtonClick;
            _view.OnAddTestMethodButtonClick += OnExplorerAddTestMethodButtonClick;
            _view.OnAddTestModuleButtonClick += OnExplorerAddTestModuleButtonClick;

            _testEngine.TestComplete += TestComplete;
        }
    }
}
