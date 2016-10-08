using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Presenter for the to-do items explorer.
    /// </summary>
    public class ToDoExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public ToDoExplorerDockablePresenter(IVBE vbe, IAddIn addin, ToDoExplorerWindow window)
            : base(vbe, addin, window)
        {
        }
    }
}
