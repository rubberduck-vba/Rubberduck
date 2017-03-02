using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    public interface IAssignedByValParameterQuickFixDialog : IDialogView
    {
        DialogResult DialogResult { get;}
        string NewName { get; set; }
    }
}
