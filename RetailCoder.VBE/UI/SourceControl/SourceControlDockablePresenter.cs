using System.Diagnostics;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Presenter for the source control view.
    /// </summary>
    public class SourceControlDockablePresenter : DockableToolwindowPresenter
    {
        public SourceControlDockablePresenter(IVBE vbe, IAddIn addin, SourceControlPanel window, IConfigProvider<WindowSettings> settings)
            : base(vbe, addin, window, settings)
        {
        }

        public SourceControlPanel Window()
        {
            var control = UserControl as SourceControlPanel;
            Debug.Assert(control != null);
            return control;
        }
    }
}
