using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerDockablePresenter : DockablePresenterBase
    {
        public TestExplorerDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl control)
            : base(vbe, addin, control)
        {
            
        }
    }
}
