using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObsoleteLetStatementInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;

        public ObsoleteLetStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            if (_parseTreeResults == null)
            {
                return Enumerable.Empty<IInspectionResult>();
            }
            return _parseTreeResults.Where(context => context.Context is VBAParser.LetStmtContext 
                && !IsIgnoringInspectionResultFor(context.ModuleName.Component, context.Context.Start.Line))
                .Select(context => new ObsoleteLetStatementUsageInspectionResult(this, new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }

        public class ObsoleteLetStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.LetStmtContext> _contexts = new List<VBAParser.LetStmtContext>();
            public IEnumerable<VBAParser.LetStmtContext> Contexts => _contexts;

            public override void ExitLetStmt(VBAParser.LetStmtContext context)
            {
                if (context.LET() != null)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
