using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Castle.Components.DictionaryAdapter;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteLetStatementInspection : IInspection
    {
        public ObsoleteLetStatementInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "ObsoleteLetStatementInspection"; } }
        public string Description { get { return RubberduckUI.ObsoleteLet; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = new List<ObsoleteLetStatementUsageInspectionResult>();
            foreach (var result in parseResult.ComponentParseResults)
            {
                var listener = new ObsoleteLetStatementListener();
                var walker = new ParseTreeWalker();

                walker.Walk(listener, result.ParseTree);
                issues.AddRange(listener.Contexts.Select(context => 
                    new ObsoleteLetStatementUsageInspectionResult(this, new QualifiedContext<ParserRuleContext>(result.QualifiedName, context))));
            }

            return issues;
        }

        private class ObsoleteLetStatementListener : VBABaseListener
        {
            private readonly IList<VBAParser.LetStmtContext> _contexts = new EditableList<VBAParser.LetStmtContext>();
            public IEnumerable<VBAParser.LetStmtContext> Contexts { get { return _contexts; } }

            public override void EnterLetStmt(VBAParser.LetStmtContext context)
            {
                if (context.LET() != null)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}