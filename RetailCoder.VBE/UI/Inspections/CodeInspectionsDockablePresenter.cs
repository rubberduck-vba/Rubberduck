using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

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
