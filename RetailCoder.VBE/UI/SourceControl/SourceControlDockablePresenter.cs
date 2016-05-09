using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Presenter for the source control view.
    /// </summary>
    public class SourceControlDockablePresenter : DockableToolwindowPresenter
    {

        public SourceControlDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl window)
            : base(vbe, addin, window)
        {
        }

        public UserControl Window()
        {
            return UserControl;
        }
    }
}
