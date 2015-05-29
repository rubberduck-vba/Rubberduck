using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Inspections;

namespace Rubberduck.UI.CodeInspections
{
    public interface ICodeInspectionsWindow : IDockableUserControl
    {
        DataGridView GridView { get; }
        event EventHandler<DataGridViewCellMouseEventArgs> SortColumn;
        event EventHandler RefreshCodeInspections;
        event EventHandler<NavigateCodeEventArgs> NavigateCodeIssue;
        event EventHandler<QuickFixEventArgs> QuickFix;
        BindingList<CodeInspectionResultGridViewItem> InspectionResults { get; set; }
        void ToggleParsingStatus(bool enabled = true);
        void SetIssuesStatus(int issueCount, bool completed = false);
        void SetContent(IEnumerable<CodeInspectionResultGridViewItem> inspectionResults);
    }
}
