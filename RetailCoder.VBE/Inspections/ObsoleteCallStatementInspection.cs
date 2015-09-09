using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementInspection : IInspection
    {
        public ObsoleteCallStatementInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "ObsoleteCallStatementInspection"; } }
        public string Description { get { return RubberduckUI.ObsoleteCall; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = new List<ObsoleteCallStatementUsageInspectionResult>();
            foreach (var result in parseResult.ComponentParseResults)
            {
                var listener = new ObsoleteCallStatementListener();
                var walker = new ParseTreeWalker();

                walker.Walk(listener, result.ParseTree);
                issues.AddRange(listener.Contexts.Select(context => new ObsoleteCallStatementUsageInspectionResult(Description, Severity,
                    new QualifiedContext<VBAParser.ExplicitCallStmtContext>(result.QualifiedName, context))));
            }

            return issues;
        }

        private class ObsoleteCallStatementListener : VBABaseListener
        {
            private readonly IList<VBAParser.ExplicitCallStmtContext> _contexts = new List<VBAParser.ExplicitCallStmtContext>();
            public IEnumerable<VBAParser.ExplicitCallStmtContext> Contexts { get { return _contexts; } }

            public override void EnterExplicitCallStmt(VBAParser.ExplicitCallStmtContext context)
            {
                var procedureCall = context.eCS_ProcedureCall();
                if (procedureCall != null)
                {
                    if (procedureCall.CALL() != null)
                    {
                        _contexts.Add(context);
                        return;
                    }
                }

                var memberCall = context.eCS_MemberProcedureCall();
                if (memberCall == null) return;
                if (memberCall.CALL() == null) return;
                _contexts.Add(context);
            }
        }
    }
}