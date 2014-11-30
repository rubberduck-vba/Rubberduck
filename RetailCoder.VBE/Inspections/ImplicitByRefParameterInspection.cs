using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitByRefParameterInspection : CodeInspection
    {
        public ImplicitByRefParameterInspection()
            : base("Parameter is passed ByRef implicitly",
                "Pass parameter ByRef explicitly, or pass it ByVal if it isn't assigned in this procedure.",
                CodeInspectionType.CodeQualityIssues,
                CodeInspectionSeverity.Warning)
        {
        }

        public override IEnumerable<CodeInspectionResultBase> Inspect(SyntaxTreeNode node)
        {
            var procedures = node.FindAllProcedures().Where(procedure => procedure.Parameters.Any(parameter => !string.IsNullOrEmpty(parameter.Instruction.Value)));
            var targets = procedures.Where(procedure => procedure.Parameters.Any(parameter => parameter.IsImplicitByRef));

            return targets.SelectMany(procedure => procedure.Parameters.Where(parameter => parameter.IsImplicitByRef)
                .Select(parameter => new ImplicitByRefParameterInspectionResult(Name, parameter.Instruction, Severity, QuickFixMessage)));
        }
    }
}