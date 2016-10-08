using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public TestExplorerDockablePresenter(IVBE vbe, IAddIn addin, TestExplorerWindow view)
            : base(vbe, addin, view)
        {
        }
    }
}
