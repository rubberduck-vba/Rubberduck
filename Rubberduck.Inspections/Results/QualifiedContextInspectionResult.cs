using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    internal class QualifiedContextInspectionResult : InspectionResultBase
    {
        public QualifiedContextInspectionResult(IInspection inspection, string description, QualifiedContext context, Dictionary<string, string> properties = null) :
            base(inspection,
                 description,
                 context.ModuleName,
                 context.Context,
                 null,
                 new QualifiedSelection(context.ModuleName, context.Context.GetSelection()),
                 context.MemberName,
                 properties)
        {
        }
    }
}
