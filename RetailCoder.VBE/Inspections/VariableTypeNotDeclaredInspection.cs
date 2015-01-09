using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class VariableTypeNotDeclaredInspection : IInspection
    {
        public VariableTypeNotDeclaredInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.VariableTypeNotDeclared; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(SyntaxTreeNode node)
        {
            var identifiers = node.FindAllDeclarations()
                .Where(declaration => !declaration.Instruction.Line.IsMultiline)
                .SelectMany(declaration => declaration.ChildNodes.Cast<IdentifierNode>())
                .Where(identifier => !identifier.IsTypeSpecified);

            return identifiers.Select(identifier => new VariableTypeNotDeclaredInspectionResult(Name, identifier, Severity));
        }
    }
}