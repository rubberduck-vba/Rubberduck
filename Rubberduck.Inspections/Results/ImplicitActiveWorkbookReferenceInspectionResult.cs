using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitActiveWorkbookReferenceInspectionResult : InspectionResultBase
    {
        public ImplicitActiveWorkbookReferenceInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitActiveWorkbookReferenceInspectionResultFormat, Context.GetText()); }
        }
    }
}