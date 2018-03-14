using System;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    public interface IDialogView : IDisposable
    {
        DialogResult ShowDialog();
    }
}
