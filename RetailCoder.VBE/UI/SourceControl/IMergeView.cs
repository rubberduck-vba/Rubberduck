using System;
using System.Collections.Generic;

namespace Rubberduck.UI.SourceControl
{
    public enum MergeStatus { Unknown, Success, Failure}

    public interface IMergeView
    {
        bool OkButtonEnabled { get; set; }
        IList<string> SourceSelectorData { get; set; }
        IList<string> DestinationSelectorData { get; set; }
        string SelectedSourceBranch { get; set; }
        string SelectedDestinationBranch { get; set; }
        MergeStatus Status { get; set; }
        string StatusText { get; set; }
        bool StatusTextVisible { get; set; }
        
        event EventHandler<EventArgs> Confirm;
        event EventHandler<EventArgs> Cancel;
        event EventHandler<EventArgs> SelectedSourceBranchChanged;
        event EventHandler<EventArgs> SelectedDestinationBranchChanged;
        event EventHandler<EventArgs> MergeStatusChanged;

        //cherry pick Form methods/properties to expose
        void Show();
        void Hide();
        void Close();
    }
}
