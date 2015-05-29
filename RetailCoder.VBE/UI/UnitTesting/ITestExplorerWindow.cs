using Rubberduck.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.UnitTesting
{
    public interface ITestExplorerWindow : IDockableUserControl
    {
        DataGridView GridView { get; }
        event EventHandler<DataGridViewCellMouseEventArgs> SortColumn;
        BindingList<TestExplorerItem> AllTests { get; set; }
        string ClassId { get; }
        string Caption { get; }
        event EventHandler<SelectedTestEventArgs> OnRunSelectedTestButtonClick;
        event EventHandler OnRunLastRunTestsButtonClick;
        event EventHandler OnRunNotRunTestsButtonClick;
        event EventHandler OnRunPassedTestsButtonClick;
        event EventHandler OnRunFailedTestsButtonClick;
        event EventHandler OnRunAllTestsButtonClick;
        event EventHandler OnAddExpectedErrorTestMethodButtonClick;
        event EventHandler OnAddTestMethodButtonClick;
        event EventHandler OnAddTestModuleButtonClick;
        void ClearProgress();
        void ClearResults();
        void Refresh(IDictionary<TestMethod, TestResult> tests);
        void SetPlayList(IEnumerable<TestMethod> tests);
        void SetPlayList(IDictionary<TestMethod, TestResult> tests);
        event EventHandler OnRefreshListButtonClick;
        event EventHandler<SelectedTestEventArgs> OnGoToSelectedTest;
        void WriteResult(TestMethod test, TestResult result);
    }
}
