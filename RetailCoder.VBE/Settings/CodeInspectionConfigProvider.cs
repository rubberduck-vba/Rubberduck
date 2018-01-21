using System.Collections.Generic;
using System.Linq;
using Rubberduck.SettingsProvider;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Settings
{
    public class CodeInspectionConfigProvider : IConfigProvider<CodeInspectionSettings>
    {
        private readonly IPersistanceService<CodeInspectionSettings> _persister;
        private readonly CodeInspectionSettings _defaultSettings;

        public CodeInspectionConfigProvider(IPersistanceService<CodeInspectionSettings> persister, IEnumerable<IInspection> foundInspections)
        {
            _persister = persister;
            _defaultSettings = new DefaultSettings<CodeInspectionSettings>().Default;

            var nonDefaultInspections = foundInspections
                .Where(inspection => !_defaultSettings.CodeInspections.Select(x => x.Name).Contains(inspection.Name));

            _defaultSettings.CodeInspections.UnionWith(nonDefaultInspections.Select(inspection => new CodeInspectionSetting(inspection)));
        }

        public CodeInspectionSettings Create()
        {
            // Loaded settings don't contain defaults, so we need to combine user settings with defaults.
            var loaded = _persister.Load(_defaultSettings);
            loaded?.CodeInspections.UnionWith(_defaultSettings.CodeInspections);

            return loaded ?? _defaultSettings;
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
