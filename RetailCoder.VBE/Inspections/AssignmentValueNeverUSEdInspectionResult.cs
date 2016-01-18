using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class AssignmentValueNeverUsedInspectionResult : CodeInspectionResultBase
    {
        public AssignmentValueNeverUsedInspectionResult(IInspection inspection, IdentifierReference target)
            : base(inspection, string.Format(inspection.Description, target.IdentifierName), target.QualifiedModuleName, target.Context)
        {
        }
    }
}
