using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteLetStatementInspection : InspectionBase, IParseTreeInspection<VBAParser.LetStmtContext>
    {
        private IEnumerable<QualifiedContext> _results;

        public ObsoleteLetStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ObsoleteLetStatementInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteLetStatementInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public IEnumerable<QualifiedContext<VBAParser.LetStmtContext>> ParseTreeResults { get { return _results.OfType<QualifiedContext<VBAParser.LetStmtContext>>(); } }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _results = results;
        }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }
            return ParseTreeResults.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName.Component, context.Context.Start.Line))
                .Select(context => new ObsoleteLetStatementUsageInspectionResult(this, new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }

        public class ObsoleteLetStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.LetStmtContext> _contexts = new List<VBAParser.LetStmtContext>();
            public IEnumerable<VBAParser.LetStmtContext> Contexts { get { return _contexts; } }

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
