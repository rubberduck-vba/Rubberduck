using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public CodeExplorerDockablePresenter(IVBE vbe, IAddIn addIn, CodeExplorerWindow view)
            : base(vbe, addIn, view)
        {
        }
    }
}
