using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.SourceControl
{
    public interface IBranchesView
    {
        IList<string> Branches { get; set; }
        string CurrentBranch { get; set; }
        IList<string> PublishedBranches { get; set; }
        IList<string> UnpublishedBranches { get; set; } 

        event EventHandler<EventArgs> SelectedBranchChanged;
        event EventHandler<EventArgs> Merge;
        event EventHandler<EventArgs> CreateBranch;
        event EventHandler<EventArgs> RefreshData;
    }
}
