using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsDockablePresenter : DockableToolwindowPresenter
    {
        public CodeInspectionsDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl window)
            :base(vbe, addin, window)
        {
        }
    }
}
