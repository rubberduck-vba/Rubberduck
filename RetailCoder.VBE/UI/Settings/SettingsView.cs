using System.Globalization;

namespace Rubberduck.UI.Settings
{
    public class SettingsView
    {
        public string Label { get { return RubberduckUI.ResourceManager.GetString("SettingsCaption_" + View); } }
        public string Instructions
        {
            get
            {
                return RubberduckUI.ResourceManager.GetString("SettingsInstructions_" + View, CultureInfo.CurrentUICulture);
            }
        }
        public ISettingsView Control { get; set; }
        public SettingsViews View { get; set; }
    }
}
