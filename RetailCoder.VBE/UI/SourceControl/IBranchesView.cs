using System;
using System.Collections.Generic;

namespace Rubberduck.UI.SourceControl
{
    public interface IBranchesView
    {
        IList<string> Local { get; set; }
        string Current { get; set; }
        IList<string> Published { get; set; }
        IList<string> Unpublished { get; set; }

        event EventHandler<EventArgs> SelectedBranchChanged;
        event EventHandler<EventArgs> Merge;
        event EventHandler<EventArgs> CreateBranch;
        event EventHandler<EventArgs> DeleteBranch;
    }
}
