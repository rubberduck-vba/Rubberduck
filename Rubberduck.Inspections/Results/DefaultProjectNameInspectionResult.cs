using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class DefaultProjectNameInspectionResult : InspectionResultBase
    {
        public DefaultProjectNameInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target) {}

        public override string Description
        {
            get { return Inspection.Description.Capitalize(); }
        }
    }
}
