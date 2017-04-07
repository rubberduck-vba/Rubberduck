using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ParameterCanBeByValInspectionResult : InspectionResultBase
    {
        public ParameterCanBeByValInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target.QualifiedName.QualifiedModuleName, target.Context, target) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterCanBeByValInspectionResultFormat, Target.IdentifierName).Capitalize(); }
        }
    }
}
