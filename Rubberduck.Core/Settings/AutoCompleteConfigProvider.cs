using System.Collections.Generic;
using System.Linq;
using Rubberduck.AutoComplete;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
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

            var defaultKeys = _defaultSettings.AutoCompletes.Select(e => e.Key);
            var nonDefaultAutoCompletes = provider.AutoCompletes.Where(e => !defaultKeys.Contains(e.GetType().Name));

            _defaultSettings.AutoCompletes.UnionWith(nonDefaultAutoCompletes.Select(e => new AutoCompleteSetting(e)));
        }

        public AutoCompleteSettings Create()
        {
            var loaded = _persister.Load(_defaultSettings);
            if (loaded == null)
            {
                return _defaultSettings;
            }

            // Loaded settings don't contain defaults, so we need to combine user settings with defaults.
            var settings = new HashSet<AutoCompleteSetting>();

            foreach (var loadedSetting in loaded.AutoCompletes.Where(e => !settings.Contains(e) && _foundAutoCompleteKeys.Contains(e.Key)))
            {
                var matchingDefaultSetting = _defaultSettings.AutoCompletes.FirstOrDefault(e => !loaded.AutoCompletes.Contains(e) && e.Equals(loadedSetting));
                if (matchingDefaultSetting != null)
                {
                    loadedSetting.IsEnabled = matchingDefaultSetting.IsEnabled;
                }

                settings.Add(loadedSetting);
            }

            settings.UnionWith(_defaultSettings.AutoCompletes.Where(e => !settings.Contains(e)));
            loaded.AutoCompletes = settings;

            return loaded;
        }

        public AutoCompleteSettings CreateDefaults()
        {
            return _defaultSettings;
        }

        public void Save(AutoCompleteSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
