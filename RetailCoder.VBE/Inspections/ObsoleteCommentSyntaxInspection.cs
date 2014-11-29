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
            _name = "Use of obsolete Rem comment syntax";
            _quickFixMessage = "Replace Rem reserved keyword with single quote.";
            _inspectionType = CodeInspectionType.MaintainabilityAndReadabilityIssues;
            _severity = CodeInspectionSeverity.Suggestion;
        }

        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly string _quickFixMessage;
        public string QuickFixMessage { get { return _quickFixMessage; } }

        private readonly CodeInspectionType _inspectionType;
        public CodeInspectionType InspectionType { get { return _inspectionType; } }

        private readonly CodeInspectionSeverity _severity;
        public CodeInspectionSeverity Severity { get { return _severity; } }

        public bool IsEnabled { get; set; }

        public IEnumerable<CodeInspectionResultBase> Inspect(SyntaxTreeNode node)
        {
            return node.FindAllComments()
                       .Where(instruction => instruction.Value == ReservedKeywords.Rem)
                       .Select(instruction => new ObsoleteCommentSyntaxInspectionResult(instruction, _severity, _quickFixMessage));
        }
    }
}