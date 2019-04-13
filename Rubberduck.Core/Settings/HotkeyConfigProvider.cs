using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class HotkeyConfigProvider : ConfigurationServiceBase<HotkeySettings>
    {
        private readonly IEnumerable<HotkeySetting> _defaultHotkeys;

        public HotkeyConfigProvider(IPersistanceService<HotkeySettings> persister)
            : base(persister)
        {
            _defaultHotkeys = new DefaultSettings<HotkeySetting, Properties.Settings>().Defaults;
        }

        public override HotkeySettings Load()
        {
            var prototype = new HotkeySettings(_defaultHotkeys);

            // Loaded settings don't contain defaults, so we need to use the `Settings` property to combine user settings with defaults.
            var loaded = persister.Load(prototype);
            if (loaded != null)
            {
                prototype.Settings = loaded.Settings;
            }

            return prototype;
        }

        public override HotkeySettings LoadDefaults()
        {
            return new HotkeySettings(_defaultHotkeys);
        }
    }
}
