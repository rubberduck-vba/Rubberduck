using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class EmptyStringLiteralInspectionResult : InspectionResultBase
    {
        public EmptyStringLiteralInspectionResult(IInspection inspection, QualifiedContext qualifiedContext, QualifiedMemberName? qualifiedName)
            : base(inspection, qualifiedContext.ModuleName, qualifiedName, qualifiedContext.Context) {}

        public override string Description
        {
            get { return InspectionsUI.EmptyStringLiteralInspectionResultFormat.Capitalize(); }
        }
    }
}
