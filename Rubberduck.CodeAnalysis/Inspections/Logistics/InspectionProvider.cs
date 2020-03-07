using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Settings;

namespace Rubberduck.CodeAnalysis.Inspections.Logistics
{
    internal class InspectionProvider : IInspectionProvider
    {
        public InspectionProvider(IEnumerable<IInspection> inspections)
        {
            var defaultSettings = new DefaultSettings<CodeInspectionSettings, Properties.CodeInspectionDefaults>().Default;
            var defaultNames = defaultSettings.CodeInspections.Select(x => x.Name);
            var defaultInspections = inspections.Where(inspection => defaultNames.Contains(inspection.Name));

            foreach (var inspection in defaultInspections)
            {
                inspection.InspectionType = defaultSettings.CodeInspections.First(setting => setting.Name == inspection.Name).InspectionType;
            }
            
            Inspections = inspections;
        }

        public IEnumerable<IInspection> Inspections { get; }
    }
}
