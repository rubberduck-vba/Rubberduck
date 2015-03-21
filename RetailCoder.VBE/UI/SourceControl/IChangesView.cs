using System;
using System.Collections.Generic;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Specifies what actions to take on Commit.
    /// </summary>
    public enum CommitAction { Commit, CommitAndPush, CommitAndSync }

    /// <summary>
    /// Defines a view of changes to be committed.
    /// </summary>
    public interface IChangesView
    {
        string CommitMessage { get; set; }
        CommitAction CommitAction { get; set; }
        //todo: support directories
        IList<string> IncludedChanges { get; set; }
        IList<string> ExcludedChanges { get; set; }
        IList<string> UntrackedFiles { get; set; } 

        event EventHandler<EventArgs> Commit;
        event EventHandler<EventArgs> RefreshData;
    }
}
