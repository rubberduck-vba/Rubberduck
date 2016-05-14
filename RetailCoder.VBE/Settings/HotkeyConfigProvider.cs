using System.Collections.Generic;
using System.Windows.Input;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public interface IHotkeyConfigProvider
    {
        HotkeySettings Create();
        HotkeySettings CreateDefaults();
        void Save(HotkeySettings settings);
    }

    public class HotkeyConfigProvider : IHotkeyConfigProvider
    {
        private readonly IPersistanceService<HotkeySettings> _persister;

        public HotkeyConfigProvider(IPersistanceService<HotkeySettings> persister)
        {
            _persister = persister;
            
        }

        public HotkeySettings Create()
        {          
            var prototype = new HotkeySettings();
            return _persister.Load(prototype) ?? prototype;
        }

        public HotkeySettings CreateDefaults()
        {
            return new HotkeySettings();
        }

        public void Save(HotkeySettings settings)
        {
            _persister.Save(settings);
        }
    }
}
