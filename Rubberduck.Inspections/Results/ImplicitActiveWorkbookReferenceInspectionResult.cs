using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitActiveWorkbookReferenceInspectionResult : InspectionResultBase
    {
        public ImplicitActiveWorkbookReferenceInspectionResult(IInspection inspection, IdentifierReference reference, QualifiedMemberName? qualifiedName)
            : base(inspection, reference.QualifiedModuleName, qualifiedName, reference.Context) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitActiveWorkbookReferenceInspectionResultFormat, Context.GetText()); }
        }
    }
}