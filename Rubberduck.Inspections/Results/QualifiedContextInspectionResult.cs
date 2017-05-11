using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    class QualifiedContextInspectionResult : InspectionResultBase
    {
        public QualifiedContextInspectionResult(IInspection inspection, string description, RubberduckParserState state, QualifiedContext context, Dictionary<string, string> properties = null) :
            base(inspection,
                 description,
                 context.ModuleName,
                 context.Context,
                 null,
                 new QualifiedSelection(context.ModuleName, context.Context.GetSelection()),
                 GetQualifiedMemberName(state, context),
                 properties)
        {
        }

        private static QualifiedMemberName? GetQualifiedMemberName(RubberduckParserState state, QualifiedContext context)
        {
            var members = state.DeclarationFinder.Members(context.ModuleName);
            return members.SingleOrDefault(m => m.Selection.Contains(context.Context.GetSelection()))?.QualifiedName;
        }
    }
}
