using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteCommentSyntaxInspection : InspectionBase
    {
        /// <summary>
        /// Parameterless constructor required for discovery of implemented code inspections.
        /// </summary>
        public ObsoleteCommentSyntaxInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public override string Description { get { return RubberduckUI.ObsoleteComment; } }
        public override CodeInspectionType InspectionType { get {return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            return State.AllComments.Where(comment => comment.Marker == Tokens.Rem)
                .Select(comment => new ObsoleteCommentSyntaxInspectionResult(this, comment));
        }
    }
}