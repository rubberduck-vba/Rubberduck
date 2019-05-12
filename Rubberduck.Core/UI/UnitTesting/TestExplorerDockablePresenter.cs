using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting
{
    internal class TestExplorerDockablePresenter : DockableToolwindowPresenter
    {
        public TestExplorerDockablePresenter(IVBE vbe, IAddIn addin, TestExplorerWindow view, IConfigurationService<WindowSettings> settings)
            : base(vbe, addin, view, settings)
        {
        }
    }
}
