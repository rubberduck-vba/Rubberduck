using Rubberduck.Resources;

namespace Rubberduck.UI.Settings
{
    class GeneralSettingsView : SettingsView
    {
        public override string Instructions => string.Format(base.Instructions, 
            RubberduckUI.SaveAndClose, 
            RubberduckUI.CancelButtonText);
    }
}
