using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.SettingsProvider;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;

namespace Rubberduck.CodeAnalysis.Settings
{
    public class CodeInspectionConfigProvider : ConfigurationServiceBase<CodeInspectionSettings>
    {
        private readonly CodeInspectionSettings defaultSettings;
        private readonly HashSet<string> foundInspectionNames;

        public CodeInspectionConfigProvider(IPersistanceService<CodeInspectionSettings> persister, IInspectionProvider inspectionProvider)
            : base(persister)
        {
            foundInspectionNames = inspectionProvider.Inspections.Select(inspection => inspection.Name).ToHashSet();
            defaultSettings = new DefaultSettings<CodeInspectionSettings, Properties.CodeInspectionDefaults>().Default;
            // Ignore settings for unknown inspections, for example when using the Experimental attribute
            defaultSettings.CodeInspections = defaultSettings.CodeInspections.Where(setting => foundInspectionNames.Contains(setting.Name)).ToHashSet();

            var defaultNames = defaultSettings.CodeInspections.Select(x => x.Name);
            var nonDefaultInspections = inspectionProvider.Inspections.Where(inspection => !defaultNames.Contains(inspection.Name));

            defaultSettings.CodeInspections.UnionWith(nonDefaultInspections.Select(inspection => new CodeInspectionSetting(inspection)));
        }

        public override CodeInspectionSettings Load()
        {
            var loaded = persister.Load(defaultSettings);

            if (loaded == null)
            {
                return defaultSettings;
            }

            // Loaded settings don't contain defaults, so we need to combine user settings with defaults.
            var settings = new HashSet<CodeInspectionSetting>();

            foreach (var loadedSetting in loaded.CodeInspections.Where(inspection => foundInspectionNames.Contains(inspection.Name)))
            {
                var matchingDefaultSetting = defaultSettings.CodeInspections.FirstOrDefault(inspection => inspection.Equals(loadedSetting));
                if (matchingDefaultSetting != null)
                {
                    loadedSetting.InspectionType = matchingDefaultSetting.InspectionType;
                }

                settings.Add(loadedSetting);
            }

            settings.UnionWith(defaultSettings.CodeInspections.Where(inspection => !settings.Contains(inspection)));

            loaded.CodeInspections = settings;

            return loaded;
        }

        public override CodeInspectionSettings LoadDefaults()
        {
            return defaultSettings;
        }
    }
}
