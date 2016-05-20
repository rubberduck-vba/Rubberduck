using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteLetStatementInspection : InspectionBase, IParseTreeInspection
    {
        public ObsoleteLetStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ObsoleteLetStatementInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteLetStatementInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public ParseTreeResults ParseTreeResults { get; set; }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }
            return ParseTreeResults.ObsoleteLetContexts.Select(context =>
                new ObsoleteLetStatementUsageInspectionResult(this, new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
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
