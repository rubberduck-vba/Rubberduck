using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    internal class QualifiedContextInspectionResult : InspectionResultBase
    {
        public QualifiedContextInspectionResult(IInspection inspection, string description, QualifiedContext context, dynamic properties = null) :
            base(inspection,
                 description,
                 context.ModuleName,
                 context.Context,
                 null,
                 new QualifiedSelection(context.ModuleName, context.Context.GetSelection()),
                 context.MemberName,
                 (object)properties)
        {
        }
    }
}
