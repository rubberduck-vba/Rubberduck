using System;

namespace Rubberduck.UI.SourceControl
{
    //todo: create and implement form
    public interface ICreateBranchView
    {
        string UserInputText { get; set; }
        bool IsValidBranchName { get; set; }

        event EventHandler<BranchCreateArgs> Confirm;
        event EventHandler<EventArgs> Cancel;
        event EventHandler<EventArgs> UserInputTextChanged;

        void Show();
        void Hide();
        void Close();
    }

    public class BranchCreateArgs : EventArgs
    {
        public string BranchName { get; private set; }

        public BranchCreateArgs(string branchName)
        {
            this.BranchName = branchName;
        }
    }
}
