using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class WriteOnlyPropertyInspectionResult : InspectionResultBase
    {
        public WriteOnlyPropertyInspectionResult(IInspection inspection, Declaration target) 
            : base(inspection, target) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.WriteOnlyPropertyInspectionResultFormat, Target.IdentifierName).Capitalize(); }
        }
    }
}