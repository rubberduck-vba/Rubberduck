using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class HotkeyConfigProvider : IConfigProvider<HotkeySettings>
    {
        private readonly IPersistanceService<HotkeySettings> _persister;
        private readonly IEnumerable<HotkeySetting> _defaultHotkeys;

        public HotkeyConfigProvider(IPersistanceService<HotkeySettings> persister, DefaultHotkeys defaultHotkeys)
        {
            _persister = persister;
            _defaultHotkeys = defaultHotkeys.Hotkeys;
            //_hotkeySettings = hotkeySettings;
        }

        public HotkeySettings Create()
        {
            var prototype = new HotkeySettings(_defaultHotkeys);
            return _persister.Load(prototype) ?? prototype;
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
