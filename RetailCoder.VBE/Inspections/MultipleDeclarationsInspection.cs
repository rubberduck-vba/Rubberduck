using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class MultipleDeclarationsInspection : IInspection
    {
        public MultipleDeclarationsInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.MultipleDeclarations; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(SyntaxTreeNode node)
        {
            return node.FindAllDeclarations()
                .Where(declaration => declaration.ChildNodes.Count() > 1)
                .Select(declaration => new MultipleDeclarationsInspectionResult(Name, declaration, Severity)); 
        }
    }
}