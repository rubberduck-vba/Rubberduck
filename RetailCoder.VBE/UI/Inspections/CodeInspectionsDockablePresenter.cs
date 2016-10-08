using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.UI.Inspections
{
    public class CodeInspectionsDockablePresenter : DockableToolwindowPresenter
    {
        public CodeInspectionsDockablePresenter(IVBE vbe, IAddIn addin, CodeInspectionsWindow window)
            :base(vbe, addin, window)
        {
        }
    }
}
