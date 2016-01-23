using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class SelfAssignedDeclarationInspectionResult : CodeInspectionResultBase
    {
        public SelfAssignedDeclarationInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, string.Format(inspection.Description, target.IdentifierName), target)
        {
        }
    }
}
