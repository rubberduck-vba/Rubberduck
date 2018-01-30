using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    public interface IAssignedByValParameterQuickFixDialog : IDialogView
    {
        DialogResult DialogResult { get;}
        string NewName { get; set; }
    }
}
