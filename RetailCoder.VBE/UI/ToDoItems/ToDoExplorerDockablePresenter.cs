using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Presenter for the to-do items explorer.
    /// </summary>
    public class ToDoExplorerDockablePresenter : DockableToolwindowPresenter
    {

        public ToDoExplorerDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl window)
            : base(vbe, addin, window)
        {
        }
    }
}
