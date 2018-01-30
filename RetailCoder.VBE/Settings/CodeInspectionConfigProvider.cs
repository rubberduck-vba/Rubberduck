using System.Collections.Generic;
using System.Linq;
using Rubberduck.SettingsProvider;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Settings
{
    public class CodeInspectionConfigProvider : IConfigProvider<CodeInspectionSettings>
    {
        private readonly IPersistanceService<CodeInspectionSettings> _persister;
        private readonly CodeInspectionSettings _defaultSettings;
        private readonly HashSet<string> _foundInspectionNames;

        public CodeInspectionConfigProvider(IPersistanceService<CodeInspectionSettings> persister, IEnumerable<IInspection> foundInspections)
        {
            _persister = persister;
            _foundInspectionNames = foundInspections.Select(inspection => inspection.Name).ToHashSet();
            _defaultSettings = new DefaultSettings<CodeInspectionSettings>().Default;

            var defaultNames = _defaultSettings.CodeInspections.Select(x => x.Name).ToHashSet();

            var defaultInspections = foundInspections.Where(inspection => defaultNames.Contains(inspection.Name));
            var nonDefaultInspections = foundInspections.Except(defaultInspections);

            foreach (var inspection in defaultInspections)
            {
                inspection.InspectionType = _defaultSettings.CodeInspections.First(setting => setting.Name == inspection.Name).InspectionType;
            }

            _defaultSettings.CodeInspections.UnionWith(nonDefaultInspections.Select(inspection => new CodeInspectionSetting(inspection)));
        }

        public CodeInspectionSettings Create()
        {
            var loaded = _persister.Load(_defaultSettings);

            if (loaded == null)
            {
                return _defaultSettings;
            }

            var settings = new HashSet<CodeInspectionSetting>();

            // Loaded settings don't contain defaults, so we need to combine user settings with defaults.
            foreach (var loadedSetting in loaded.CodeInspections.Where(inspection => _foundInspectionNames.Contains(inspection.Name)).Distinct())
            {
                var matchingDefaultSetting = _defaultSettings.CodeInspections.FirstOrDefault(inspection => inspection.Equals(loadedSetting));
                if (matchingDefaultSetting != null)
                {
                    loadedSetting.InspectionType = matchingDefaultSetting.InspectionType;
                }

                settings.Add(loadedSetting);
            }

            settings.UnionWith(_defaultSettings.CodeInspections.Where(inspection => !settings.Contains(inspection)));

            loaded.CodeInspections = settings;

            return loaded;
        }

        public CodeInspectionSettings CreateDefaults()
        {
            return _defaultSettings;
        }

        public void Save(CodeInspectionSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
