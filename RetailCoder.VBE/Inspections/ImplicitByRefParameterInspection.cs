using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
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
        
        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IDictionary<QualifiedModuleName,IParseTree> nodes)
        {
            var signatures = nodes.SelectMany(node =>
                node.Value.GetPublicProcedures()
                    .Select(procedure => new
                    {
                        QualifiedName = node.Key,
                        ProcedureName = procedure.ambiguousIdentifier().GetText(),
                        Parameters = procedure.argList()
                    }));

            var targets =
                signatures.SelectMany(
                    signature => signature.Parameters.arg().Where(arg => arg.BYREF() == null && arg.BYVAL() == null)
                        .Select(arg => new {signature.QualifiedName, signature.ProcedureName, arg}));

            return targets.Select(parameter => new ImplicitByRefParameterInspectionResult(Name, parameter.arg, Severity, parameter.QualifiedName.ProjectName, parameter.QualifiedName.ModuleName, parameter.ProcedureName));
        }
    }
}