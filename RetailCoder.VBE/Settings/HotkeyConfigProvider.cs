using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class HotkeyConfigProvider : IConfigProvider<HotkeySettings>
    {
        private readonly IPersistanceService<HotkeySettings> _persister;
        //private readonly HotkeySettings _hotkeySettings;
        private readonly IEnumerable<HotkeySetting> _defaultHotkeys;

        //public IEnumerable<HotkeySetting> DefaultHotkeys { get; set; }

        //public HotkeyConfigProvider(IPersistanceService<HotkeySettings> persister)
        //public HotkeyConfigProvider(IPersistanceService<HotkeySettings> persister, IEnumerable<CommandBase> commands)
        public HotkeyConfigProvider(IPersistanceService<HotkeySettings> persister, IEnumerable<HotkeySetting> defaultHotkeys)
        //public HotkeyConfigProvider(IPersistanceService<HotkeySettings> persister, HotkeySettings hotkeySettings)
        {
            _persister = persister;
            _defaultHotkeys = defaultHotkeys;
            //_hotkeySettings = hotkeySettings;
        }

        public HotkeySettings Create()
        {
            var prototype = new HotkeySettings(_defaultHotkeys);
            //var prototype = new HotkeySettings(DefaultHotkeys);
            return _persister.Load(prototype) ?? prototype;
        }

        public HotkeySettings CreateDefaults()
        {
            return new HotkeySettings(_defaultHotkeys);
            //return new HotkeySettings(DefaultHotkeys);
        }

        public void Save(HotkeySettings settings)
        {
            _persister.Save(settings);
        }
    }
}
