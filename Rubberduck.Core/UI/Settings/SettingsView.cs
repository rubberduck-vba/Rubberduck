using System.Globalization;
using Rubberduck.Resources.Settings;

namespace Rubberduck.UI.Settings
{
    public class SettingsView
    {
        public string Label => SettingsUI.ResourceManager.GetString("PageHeader_" + View);
        public virtual string Instructions => SettingsUI.ResourceManager.GetString("PageInstructions_" + View, CultureInfo.CurrentUICulture);
        public ISettingsView Control { get; set; }
        public SettingsViews View { get; set; }
    }
}
