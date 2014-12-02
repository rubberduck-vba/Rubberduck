using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ObsoleteCommentSyntaxInspection : IInspection
    {
        /// <summary>
        /// Parameterless constructor required for discovery of implemented code inspections.
        /// </summary>
        public ObsoleteCommentSyntaxInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "Use of obsolete Rem comment syntax"; } }
        public CodeInspectionType InspectionType { get {return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(SyntaxTreeNode node)
        {
            var comments = node.FindAllComments();
            var remComments = comments.Where(instruction => instruction.Comment.StartsWith(ReservedKeywords.Rem));
            return remComments.Select(instruction => new ObsoleteCommentSyntaxInspectionResult(Name, instruction, Severity));
        }
    }
}