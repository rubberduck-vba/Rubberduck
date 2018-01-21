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
            return _persister.Load(_defaultSettings) ?? _defaultSettings;
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
