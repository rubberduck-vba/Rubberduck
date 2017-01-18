using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class HostSpecificExpressionInspectionResult : InspectionResultBase
    {
        public HostSpecificExpressionInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.HostSpecificExpressionInspectionResultFormat, Target.IdentifierName);
            }
        }
    }
}