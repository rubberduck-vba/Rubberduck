using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public TestExplorerDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl view)
            : base(vbe, addin, view)
        {
        }
    }
}
