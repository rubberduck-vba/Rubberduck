using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementUsageInspection : IInspection
    {
        public ObsoleteCallStatementUsageInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ObsoleteCall; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            foreach (var result in parseResult)
            {
                var statements = result.ParseTree.GetObsoleteStatements().ToList();
                var module = result;
                foreach (var inspectionResult in 
                    statements.OfType<VisualBasic6Parser.ECS_MemberProcedureCallContext>()
                              .Where(call => call.CALL() != null && !string.IsNullOrEmpty(call.CALL().GetText())).Select(node => node.Parent).Union(statements.OfType<VisualBasic6Parser.ECS_ProcedureCallContext>().Where(call => call.CALL() != null && !string.IsNullOrEmpty(call.CALL().GetText()))
                              .Select(node => node.Parent))
                              .Cast<VisualBasic6Parser.ExplicitCallStmtContext>()
                              .Select(context => 
                                  new ObsoleteCallStatementUsageInspectionResult(Name, Severity, 
                                      new QualifiedContext<VisualBasic6Parser.ExplicitCallStmtContext>(module.QualifiedName, context))))
                    yield return inspectionResult;
            }
        }
    }
}