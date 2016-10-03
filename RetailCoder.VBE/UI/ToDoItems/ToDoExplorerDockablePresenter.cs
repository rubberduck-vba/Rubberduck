using Rubberduck.VBEditor.DisposableWrappers;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Presenter for the to-do items explorer.
    /// </summary>
    public class ToDoExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public ToDoExplorerDockablePresenter(VBE vbe, AddIn addin, ToDoExplorerWindow window)
            : base(vbe, addin, window)
        {
        }
    }
}
