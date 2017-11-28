using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class HotkeyConfigProvider : IConfigProvider<HotkeySettings>
    {
        private readonly IPersistanceService<HotkeySettings> _persister;
        private readonly IEnumerable<HotkeySetting> _defaultHotkeys;

        public HotkeyConfigProvider(IPersistanceService<HotkeySettings> persister)
        {
            _persister = persister;
            _defaultHotkeys = new DefaultHotkeySettings().Hotkeys;
        }

        public HotkeySettings Create()
        {
            var prototype = new HotkeySettings(_defaultHotkeys);

            var loaded = _persister.Load(prototype);
            if (loaded != null)
            {
                prototype.Settings = loaded.Settings;
            }

            return prototype;
        }

        public HotkeySettings CreateDefaults()
        {
            return new HotkeySettings(_defaultHotkeys);
        }

        public void Save(HotkeySettings settings)
        {
            _persister.Save(settings);
        }
    }
}
