using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.SettingsProvider;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Settings
{
    public class CodeInspectionConfigProvider : IConfigProvider<CodeInspectionSettings>
    {
        private readonly IPersistanceService<CodeInspectionSettings> _persister;
        private readonly CodeInspectionSettings _defaultSettings;
        private readonly HashSet<string> _foundInspectionNames;

        public CodeInspectionConfigProvider(IPersistanceService<CodeInspectionSettings> persister, IInspectionProvider inspectionProvider)
        {
            _persister = persister;
            _foundInspectionNames = inspectionProvider.Inspections.Select(inspection => inspection.Name).ToHashSet();
            _defaultSettings = new DefaultSettings<CodeInspectionSettings>().Default;
            // Ignore settings for unknown inspections, for example when using the Experimental attribute
            _defaultSettings.CodeInspections = _defaultSettings.CodeInspections.Where(setting => _foundInspectionNames.Contains(setting.Name)).ToHashSet();

            var defaultNames = _defaultSettings.CodeInspections.Select(x => x.Name);
            var nonDefaultInspections = inspectionProvider.Inspections.Where(inspection => !defaultNames.Contains(inspection.Name));

            _defaultSettings.CodeInspections.UnionWith(nonDefaultInspections.Select(inspection => new CodeInspectionSetting(inspection)));
        }

        public CodeInspectionSettings Create()
        {
            var loaded = _persister.Load(_defaultSettings);

            if (loaded == null)
            {
                return _defaultSettings;
            }

            // Loaded settings don't contain defaults, so we need to combine user settings with defaults.
            var settings = new HashSet<CodeInspectionSetting>();

            foreach (var loadedSetting in loaded.CodeInspections.Where(inspection => _foundInspectionNames.Contains(inspection.Name)))
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
