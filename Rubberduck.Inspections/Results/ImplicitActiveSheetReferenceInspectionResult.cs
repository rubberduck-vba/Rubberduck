using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitActiveSheetReferenceInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;

        public ImplicitActiveSheetReferenceInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitActiveSheetReferenceInspectionResultFormat, _reference.Declaration.IdentifierName).Capitalize(); }
        }
    }
}