using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerDockablePresenter : DockableToolwindowPresenter
    {
        private CodeExplorerWindow Control { get { return UserControl as CodeExplorerWindow; } }

        public CodeExplorerDockablePresenter(VBE vbe, AddIn addIn, CodeExplorerWindow view)
            : base(vbe, addIn, view)
        {
        }
    }
}
