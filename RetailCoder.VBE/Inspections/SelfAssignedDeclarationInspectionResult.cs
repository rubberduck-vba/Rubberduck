using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class SelfAssignedDeclarationInspectionResult : InspectionResultBase
    {
        public SelfAssignedDeclarationInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.SelfAssignedDeclarationInspectionResultFormat, Target.IdentifierName);
            }
        }
    }
}
