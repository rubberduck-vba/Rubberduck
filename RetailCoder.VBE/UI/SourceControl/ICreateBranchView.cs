using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.SourceControl
{
    //todo: create and implement form
    public interface ICreateBranchView
    {
        string UserInputText { get; set; }
        bool OkayButtonEnabled { get; set; }

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
