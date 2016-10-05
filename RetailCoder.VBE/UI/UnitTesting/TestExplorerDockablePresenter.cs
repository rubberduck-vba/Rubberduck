using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public TestExplorerDockablePresenter(VBE vbe, AddIn addin, TestExplorerWindow view)
            : base(vbe, addin, view)
        {
        }
    }
}
