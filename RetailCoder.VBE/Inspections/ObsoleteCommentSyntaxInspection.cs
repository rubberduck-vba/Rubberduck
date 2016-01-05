using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteCommentSyntaxInspection : IInspection
    {
        /// <summary>
        /// Parameterless constructor required for discovery of implemented code inspections.
        /// </summary>
        public ObsoleteCommentSyntaxInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "ObsoleteCommentSyntaxInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return RubberduckUI.ObsoleteComment; } }
        public CodeInspectionType InspectionType { get {return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            return (state.AllComments.Where(comment => comment.Marker == Tokens.Rem)
                .Select(comment => new ObsoleteCommentSyntaxInspectionResult(this, comment)));
        }
    }
}