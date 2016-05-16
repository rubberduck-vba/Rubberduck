using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsDockablePresenter : DockableToolwindowPresenter
    {
        public SearchResultsDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl view) 
            : base(vbe, addin, view)
        {
        }

        
    }
}