using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;

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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VbModuleParseResult> parseResult)
        {
            var modules = parseResult.ToList();
            foreach (var comment in modules.SelectMany(module => module.Comments))
            {
                if (comment.Marker == Tokens.Rem)
                {
                    yield return new ObsoleteCommentSyntaxInspectionResult(Name, Severity, comment);
                }
            }
        }
    }
}