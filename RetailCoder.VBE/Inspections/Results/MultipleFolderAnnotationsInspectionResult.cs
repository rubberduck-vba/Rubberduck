using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class MultipleFolderAnnotationsInspectionResult : InspectionResultBase
    {
        public MultipleFolderAnnotationsInspectionResult(IInspection inspection, Declaration target) 
            : base(inspection, target)
        {
        }

        public override string Description
        {
            get { return string.Format(Inspection.Description, Target.IdentifierName).Captialize(); }
        }
    }
}