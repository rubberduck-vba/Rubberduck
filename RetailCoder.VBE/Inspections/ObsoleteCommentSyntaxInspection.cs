using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

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

        public string Name { get { return InspectionNames.ObsoleteComment; } }
        public CodeInspectionType InspectionType { get {return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IDictionary<QualifiedModuleName, IParseTree> nodes)
        {
            var remComments =
                nodes.SelectMany(
                    node => node.Value.GetComments().Select(comment => new {Key = node.Key, Comment = comment}))
                    .Where(comment => comment.Comment.GetText().StartsWith(ReservedKeywords.Rem + " "));
 
            return remComments.Select(instruction => new ObsoleteCommentSyntaxInspectionResult(Name, instruction.Comment, Severity, instruction.Key.ProjectName, instruction.Key.ModuleName));
        }
    }
}