using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitByRefParameterInspection : IInspection
    {
        public ImplicitByRefParameterInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ImplicitByRef; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }
        
        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(SyntaxTreeNode node)
        {
            var procedures = node.FindAllProcedures().Where(procedure => procedure.Parameters.Any());
            var targets = procedures.Where(procedure => procedure.Parameters.Any(parameter => parameter.IsImplicitByRef)
                                                    && !procedure.Instruction.Line.IsMultiline);

            return targets.SelectMany(procedure => procedure.Parameters.Where(parameter => parameter.IsImplicitByRef)
                .Select(parameter => new ImplicitByRefParameterInspectionResult(Name, parameter, Severity)));
        }
    }
}