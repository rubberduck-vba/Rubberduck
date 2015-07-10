using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Reflection;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerDockablePresenter : DockablePresenterBase
    {
        private readonly GridViewSort<TestExplorerItem> _gridViewSort;
        private readonly ITestEngine _testEngine;
        private readonly ITestExplorerWindow _view;

        public TestExplorerDockablePresenter(VBE vbe, AddIn addin, ITestExplorerWindow control, ITestEngine testEngine, GridViewSort<TestExplorerItem> gridViewSort)
            : base(vbe, addin, control)
        {
            _testEngine = testEngine;
            _gridViewSort = gridViewSort;

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
            _view.Refresh(_testEngine.AllTests);
            _view.AllTests =
                new BindingList<TestExplorerItem>(
                    _gridViewSort.Sort(_view.AllTests.AsEnumerable(), _gridViewSort.ColumnName,
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
                _testEngine.AllTests = VBE.VBProjects
                                .Cast<VBProject>().Where(project => project.Protection != vbext_ProjectProtection.vbext_pp_locked)
                                .SelectMany(project => project.TestMethods())
                                .ToDictionary(test => test, test => _testEngine.AllTests.ContainsKey(test) ? _testEngine.AllTests[test] : null);

            }
            catch (ArgumentException)
            {
                MessageBox.Show(
                    RubberduckUI.TestExplorerDockablePresenter_MultipleTestsSameNameError,
                    RubberduckUI.Warning, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
        }

        public void RunTests()
        {
            RunTests(_testEngine.AllTests.Keys);
        }

        public void RunTests(IEnumerable<TestMethod> tests)
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
            RunTests(_testEngine.AllTests.Keys);
        }

        private void OnExplorerRunFailedTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.FailedTests());
        }

        private void OnExplorerRunLastRunTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.LastRunTests());
        }

        private void OnExplorerRunNotRunTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.NotRunTests());
        }

        private void OnExplorerRunPassedTestsButtonClick(object sender, EventArgs e)
        {
            RunTests(_testEngine.PassedTests());
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
                var selection = new Selection(startLine, startColumn, endLine, endColumn);
                codeModule.CodePane.SetSelection(selection);
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
