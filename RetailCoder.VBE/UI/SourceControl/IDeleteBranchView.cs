using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.SourceControl
{
    public interface IDeleteBranchView
    {
        bool OkButtonEnabled { get; set; }

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
