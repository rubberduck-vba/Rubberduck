using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeMetrics
{
    public class CodeMetricsDockablePresenter : DockableToolwindowPresenter
    {
        public CodeMetricsDockablePresenter(IVBE vbe, IAddIn addIn, CodeMetricsWindow view, IConfigProvider<WindowSettings> settings)
            : base(vbe, addIn, view, settings)
        {
        }
    }
}
