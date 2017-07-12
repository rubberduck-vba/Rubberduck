using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.SettingsProvider;
using Rubberduck.Parsing.Inspections.Abstract;
using System.Linq;

namespace Rubberduck.Settings
{
    public class CodeInspectionConfigProvider : IConfigProvider<CodeInspectionSettings>
    {
        private readonly IPersistanceService<CodeInspectionSettings> _persister;
        private readonly IEnumerable<IInspection> _foundInspections;

        public CodeInspectionConfigProvider(IPersistanceService<CodeInspectionSettings> persister, IEnumerable<IInspection> foundInspections)
        {
            _persister = persister;
            _foundInspections = foundInspections;
        }

        public CodeInspectionSettings Create()
        {
            var prototype = new CodeInspectionSettings(GetDefaultCodeInspections(), new WhitelistedIdentifierSetting[] { }, true);
            return _persister.Load(prototype) ?? prototype;
        }

        public CodeInspectionSettings CreateDefaults()
        {
            return new CodeInspectionSettings(GetDefaultCodeInspections(), new WhitelistedIdentifierSetting[] {}, true);
        }

        public void Save(CodeInspectionSettings settings)
        {
            _persister.Save(settings);
        }

        public IEnumerable<CodeInspectionSetting> GetDefaultCodeInspections()
        {
            return _foundInspections.Select(inspection => new CodeInspectionSetting(inspection));
        }
    }
}
