using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class IdentifierNotAssignedInspectionResult : InspectionResultBase
    {
        public IdentifierNotAssignedInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.VariableNotAssignedInspectionResultFormat, Target.IdentifierName).Capitalize(); }
        }
    }
}
