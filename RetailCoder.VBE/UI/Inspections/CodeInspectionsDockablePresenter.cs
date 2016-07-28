using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.Inspections
{
    public class CodeInspectionsDockablePresenter : DockableToolwindowPresenter
    {
        public CodeInspectionsDockablePresenter(VBE vbe, AddIn addin, CodeInspectionsWindow window)
            :base(vbe, addin, window)
        {
        }
    }
}
