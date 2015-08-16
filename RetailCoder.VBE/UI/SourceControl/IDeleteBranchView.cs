using System;
using System.Collections.Generic;

namespace Rubberduck.UI.SourceControl
{
    public interface IDeleteBranchView
    {
        bool OkButtonEnabled { get; set; }
        IList<string> Branches { get; set; }

        event EventHandler<BranchDeleteArgs> SelectionChanged;
        event EventHandler<BranchDeleteArgs> Confirm;
        event EventHandler<EventArgs> Cancel;

        void Show();
        void Hide();
        void Close();
    }

    public class BranchDeleteArgs : EventArgs
    {
        public string BranchName { get; private set; }

        public BranchDeleteArgs(string branchName)
        {
            this.BranchName = branchName;
        }
    }
}
