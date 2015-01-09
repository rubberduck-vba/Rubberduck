using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitVariantReturnTypeInspection : IInspection
    {
        public ImplicitVariantReturnTypeInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ImplicitVariantReturnType; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(SyntaxTreeNode node)
        {
            var procedures = node.FindAllProcedures().Where(procedure => procedure.HasReturnType);
            var targets = procedures.Where(procedure => string.IsNullOrEmpty(procedure.SpecifiedReturnType) 
                && !procedure.Instruction.Content.EndsWith(string.Concat(" ", ReservedKeywords.As, " ", ReservedKeywords.Variant))
                && !procedure.Instruction.Line.IsMultiline);

            return targets.Select(procedure => new ImplicitVariantReturnTypeInspectionResult(Name, procedure, Severity));
        }
    }
}