using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Inspections
{
    public class InspectionLocator
    {
        private readonly IEnumerable<IInspection> _inspections;

        public InspectionLocator(IEnumerable<IInspection> inspections)
        {
            _inspections = inspections;
        }

        public IInspection GetInspection<T>() where T: IInspection
        {
            return _inspections.FirstOrDefault(s => s.Type == typeof(T));
        }
    }
}