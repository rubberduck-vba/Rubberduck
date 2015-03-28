using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.SourceControl
{
    public interface IMergeView
    {
        bool OkayButtonEnabled { get; set; }
        IList<string> SourceSelectorData { get; set; }
        IList<string> DestinationSelectorData { get; set; }
        string SelectedSourceBranch { get; set; }
        string SelectedDestinationBranch { get; set; }

        event EventHandler<EventArgs> Confirm;
        event EventHandler<EventArgs> Cancel;
        event EventHandler<EventArgs> SelectedSourceBranchChanged;
        event EventHandler<EventArgs> SelectedDestinationBranchChanged;

        void Show();
        void Hide();
        void Close();
    }
}
