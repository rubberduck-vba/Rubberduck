using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public TestExplorerDockablePresenter(IVBE vbe, IAddIn addin, TestExplorerWindow view, IConfigProvider<WindowSettings> settings)
            : base(vbe, addin, view, settings)
        {
        }
    }
}
