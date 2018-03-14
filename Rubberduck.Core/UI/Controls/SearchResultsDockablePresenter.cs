using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsDockablePresenter : DockableToolwindowPresenter
    {
        public SearchResultsDockablePresenter(IVBE vbe, IAddIn addin, IDockableUserControl view)
            : base(vbe, addin, view, null)
        {
        }

        
    }
}
