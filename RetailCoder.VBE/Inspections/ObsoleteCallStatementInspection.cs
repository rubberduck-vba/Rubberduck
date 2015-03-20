using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementInspection : IInspection
    {
        public ObsoleteCallStatementInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ObsoleteCall; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            foreach (var result in parseResult.ComponentParseResults)
            {
                var statements = (result.ParseTree.GetContexts<ObsoleteInstrutionsListener, ParserRuleContext>(new ObsoleteInstrutionsListener(result.QualifiedName)))
                                        .Select(context => context.Context).ToList();
                var module = result;
                foreach (var inspectionResult in 
                    statements.OfType<VBAParser.ECS_MemberProcedureCallContext>()
                              .Where(call => call.CALL() != null && !string.IsNullOrEmpty(call.CALL().GetText())).Select(node => node.Parent).Union(statements.OfType<VBAParser.ECS_ProcedureCallContext>().Where(call => call.CALL() != null && !string.IsNullOrEmpty(call.CALL().GetText()))
                              .Select(node => node.Parent))
                              .Cast<VBAParser.ExplicitCallStmtContext>()
                              .Select(context => 
                                  new ObsoleteCallStatementUsageInspectionResult(Name, Severity, 
                                      new QualifiedContext<VBAParser.ExplicitCallStmtContext>(module.QualifiedName, context))))
                    yield return inspectionResult;
            }
        }
    }
}