using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public CodeExplorerDockablePresenter(IVBE vbe, IAddIn addIn, CodeExplorerWindow view, IConfigProvider<WindowSettings> settings)
            : base(vbe, addIn, view, settings)
        {
        }
    }
}
