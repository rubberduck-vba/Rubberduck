using System.Collections.Generic;
using System.Linq;
using Rubberduck.AutoComplete;
using Rubberduck.Parsing.VBA;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class AutoCompleteConfigProvider : IConfigProvider<AutoCompleteSettings>
    {
        private readonly IPersistanceService<AutoCompleteSettings> _persister;
        private readonly AutoCompleteSettings _defaultSettings;
        private readonly HashSet<string> _foundAutoCompleteKeys;

        public AutoCompleteConfigProvider(IPersistanceService<AutoCompleteSettings> persister, IAutoCompleteProvider provider)
        {
            _persister = persister;
            _foundAutoCompleteKeys = provider.AutoCompletes.Select(e => e.GetType().Name).ToHashSet();
            _defaultSettings = new DefaultSettings<AutoCompleteSettings>().Default;
            _defaultSettings.AutoCompletes = _defaultSettings.AutoCompletes.Where(setting => _foundAutoCompleteKeys.Contains(setting.Key)).ToHashSet();

        }

        public AutoCompleteSettings Create()
        {
            var prototype = new AutoCompleteSettings(_defaultSettings.AutoCompletes);

            // Loaded settings don't contain defaults, so we need to use the `Settings` property to combine user settings with defaults.
            var loaded = _persister.Load(prototype);
            if (loaded != null)
            {
                prototype.AutoCompletes = loaded.AutoCompletes;
            }

            return prototype;
        }

        public AutoCompleteSettings CreateDefaults()
        {
            return new AutoCompleteSettings(_defaultSettings.AutoCompletes);
        }

        public void Save(AutoCompleteSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
