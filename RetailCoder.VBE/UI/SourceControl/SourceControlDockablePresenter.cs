using System.Windows.Forms;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Presenter for the source control view.
    /// </summary>
    public class SourceControlDockablePresenter : DockableToolwindowPresenter
    {
        public SourceControlDockablePresenter(IVBE vbe, IAddIn addin, SourceControlPanel window)
            : base(vbe, addin, window)
        {
        }

        public UserControl Window()
        {
            return UserControl;
        }
    }
}
