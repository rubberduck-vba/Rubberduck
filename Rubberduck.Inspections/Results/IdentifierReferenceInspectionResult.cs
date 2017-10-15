using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    class IdentifierReferenceInspectionResult : InspectionResultBase
    {
        public IdentifierReferenceInspectionResult(IInspection inspection, string description, RubberduckParserState state, IdentifierReference reference, Dictionary<string, string> properties = null) :
            base(inspection,
                 description,
                 reference.QualifiedModuleName,
                 reference.Context,
                 reference.Declaration,
                 new QualifiedSelection(reference.QualifiedModuleName, reference.Context.GetSelection()),
                 GetQualifiedMemberName(state, reference),
                 properties)
        {
        }

        private static QualifiedMemberName? GetQualifiedMemberName(RubberduckParserState state, IdentifierReference reference)
        {
            var members = state.DeclarationFinder.Members(reference.QualifiedModuleName);
            return members.SingleOrDefault(m => m.Selection.Contains(reference.Selection))?.QualifiedName;
        }
    }
}
