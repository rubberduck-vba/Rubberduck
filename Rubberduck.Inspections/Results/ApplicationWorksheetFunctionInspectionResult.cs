using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class ApplicationWorksheetFunctionInspectionResult : InspectionResultBase
    {
        private readonly string _memberName;

        public ApplicationWorksheetFunctionInspectionResult(IInspection inspection, QualifiedSelection qualifiedSelection, string memberName)
            : base(inspection, qualifiedSelection.QualifiedName)
        {
            _memberName = memberName;
            QualifiedSelection = qualifiedSelection;
        }

        public override QualifiedSelection QualifiedSelection { get; }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ApplicationWorksheetFunctionInspectionResultFormat, _memberName).Capitalize(); }
        }
    }
}
