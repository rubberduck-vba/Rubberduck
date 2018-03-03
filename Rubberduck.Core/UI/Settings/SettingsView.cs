using System.Globalization;

namespace Rubberduck.UI.Settings
{
    public class SettingsView
    {
        public string Label => RubberduckUI.ResourceManager.GetString("SettingsCaption_" + View);
        public string Instructions => RubberduckUI.ResourceManager.GetString("SettingsInstructions_" + View, CultureInfo.CurrentUICulture);
        public ISettingsView Control { get; set; }
        public SettingsViews View { get; set; }
    }
}
