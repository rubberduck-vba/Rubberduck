using System;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    public interface IDialogView : IDisposable
    {
        event EventHandler CancelButtonClicked;
        void OnCancelButtonClicked();

        event EventHandler OkButtonClicked;
        void OnOkButtonClicked();

        DialogResult ShowDialog();
    }
}