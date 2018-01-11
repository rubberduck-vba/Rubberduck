using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Settings
{
    public class DefaultHotkeySettings
    {
        public IEnumerable<HotkeySetting> Hotkeys { get; }

        public DefaultHotkeySettings()
        {
            var hotkeySettingsProperties = typeof(Properties.Settings).GetProperties().Where(prop => prop.PropertyType == typeof(HotkeySetting));

            Hotkeys = hotkeySettingsProperties.Select(prop => prop.GetValue(Properties.Settings.Default)).Cast<HotkeySetting>();
        }
    }
}
