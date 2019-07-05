using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Inspections
{
    public class InspectionResultsDockablePresenter : DockableToolwindowPresenter
    {
        public InspectionResultsDockablePresenter(IVBE vbe, IAddIn addin, InspectionResultsWindow window, IConfigurationService<WindowSettings> settings)
            : base(vbe, addin, window, settings)
        {
        }
    }
}
