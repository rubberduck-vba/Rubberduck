using System;
using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    public interface IDialogView
    {
        event EventHandler CancelButtonClicked;
        void OnCancelButtonClicked();

        event EventHandler OkButtonClicked;
        void OnOkButtonClicked();

        DialogResult ShowDialog();
    }
}