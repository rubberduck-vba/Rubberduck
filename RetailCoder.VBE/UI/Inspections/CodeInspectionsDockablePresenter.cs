using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Inspections
{
    public class CodeInspectionsDockablePresenter : DockableToolwindowPresenter
    {
        public CodeInspectionsDockablePresenter(IVBE vbe, IAddIn addin, CodeInspectionsWindow window, IConfigProvider<WindowSettings> settings)
            : base(vbe, addin, window, settings)
        {
        }
    }
}
