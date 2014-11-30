using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ObsoleteCommentSyntaxInspection : CodeInspection
    {
        /// <summary>
        /// Parameterless constructor required for discovery of implemented code inspections.
        /// </summary>
        public ObsoleteCommentSyntaxInspection()
            : base("Use of obsolete Rem comment syntax", 
                   "Replace Rem reserved keyword with single quote.", 
                   CodeInspectionType.MaintainabilityAndReadabilityIssues, 
                   CodeInspectionSeverity.Suggestion)
        {
        }

        public override IEnumerable<CodeInspectionResultBase> Inspect(SyntaxTreeNode node)
        {
            var comments = node.FindAllComments();
            var remComments = comments.Where(instruction => instruction.Comment.StartsWith(ReservedKeywords.Rem));
            return remComments.Select(instruction => new ObsoleteCommentSyntaxInspectionResult(Name, instruction, Severity, QuickFixMessage));
        }
    }
}