using System;
using System.Collections.Generic;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Specifies what actions to take on Commit.
    /// </summary>
    public enum CommitAction { Unset = -1, Commit, CommitAndPush, CommitAndSync }

    /// <summary>
    /// Defines a view of changes to be committed.
    /// </summary>
    public interface IChangesView
    {
        string CommitMessage { get; set; }
        CommitAction CommitAction { get; set; }
        //todo: support directories
        IList<IFileStatusEntry> IncludedChanges { get; set; }
        IList<IFileStatusEntry> ExcludedChanges { get; set; }
        IList<IFileStatusEntry> UntrackedFiles { get; set; }
        bool CommitEnabled { get; set; }

        event EventHandler<EventArgs> SelectedActionChanged;
        event EventHandler<EventArgs> CommitMessageChanged;
        event EventHandler<EventArgs> Commit;
        event EventHandler<EventArgs> RefreshData;
    }
}
