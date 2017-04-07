using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class UnassignedVariableUsageInspectionResult : InspectionResultBase
    {
        public UnassignedVariableUsageInspectionResult(IInspection inspection, Declaration declaration)
            : base(inspection, declaration) {}

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.UnassignedVariableUsageInspectionResultFormat, Target.IdentifierName).Capitalize();
            }
        }
    }
}
