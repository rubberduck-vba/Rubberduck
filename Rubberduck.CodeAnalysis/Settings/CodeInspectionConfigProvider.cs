using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Logistics;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.SettingsProvider;
using Rubberduck.Settings;

namespace Rubberduck.CodeAnalysis.Settings
{
    internal class CodeInspectionConfigProvider : ConfigurationServiceBase<CodeInspectionSettings>
    {
        private readonly HashSet<string> foundInspectionNames;

        public CodeInspectionConfigProvider(IPersistenceService<CodeInspectionSettings> persister, IInspectionProvider inspectionProvider)
            : base(persister, new DefaultSettings<CodeInspectionSettings, Properties.CodeInspectionDefaults>())
        {
            foundInspectionNames = inspectionProvider.Inspections.Select(inspection => inspection.Name).ToHashSet();
            // Ignore settings for unknown inspections, for example when using the Experimental attribute
            Defaults.Default.CodeInspections = Defaults.Default.CodeInspections.Where(setting => foundInspectionNames.Contains(setting.Name)).ToHashSet();

            var defaultNames = Defaults.Default.CodeInspections.Select(x => x.Name);
            var nonDefaultInspections = inspectionProvider.Inspections.Where(inspection => !defaultNames.Contains(inspection.Name));

            Defaults.Default.CodeInspections.UnionWith(nonDefaultInspections.Select(inspection => new CodeInspectionSetting(inspection)));
        }

        public override CodeInspectionSettings Read()
        {
            var loaded = LoadCacheValue();
            // Loaded settings don't contain defaults, so we need to combine user settings with defaults.
            var settings = new HashSet<CodeInspectionSetting>();

            foreach (var loadedSetting in loaded.CodeInspections.Where(inspection => foundInspectionNames.Contains(inspection.Name)))
            {
                var matchingDefaultSetting = Defaults.Default.CodeInspections.FirstOrDefault(inspection => inspection.Equals(loadedSetting));
                if (matchingDefaultSetting != null)
                {
                    loadedSetting.InspectionType = matchingDefaultSetting.InspectionType;
                }

                settings.Add(loadedSetting);
            }
            settings.UnionWith(Defaults.Default.CodeInspections.Where(inspection => !settings.Contains(inspection)));

            loaded.CodeInspections = settings;
            return loaded;
        }
    }
}
