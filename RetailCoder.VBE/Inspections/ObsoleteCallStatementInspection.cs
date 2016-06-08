using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteCallStatementInspection : InspectionBase, IParseTreeInspection
    {
        public ObsoleteCallStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ObsoleteCallStatementInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteCallStatementInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public ParseTreeResults ParseTreeResults { get; set; }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }

            return ParseTreeResults.ObsoleteCallContexts.Select(context =>
                new ObsoleteCallStatementUsageInspectionResult(this,
                    new QualifiedContext<VBAParser.CallStmtContext>(context.ModuleName, context.Context as VBAParser.CallStmtContext)));
        }

        public class ObsoleteCallStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.CallStmtContext> _contexts = new List<VBAParser.CallStmtContext>();
            public IEnumerable<VBAParser.CallStmtContext> Contexts { get { return _contexts; } }

            public override void ExitCallStmt(VBAParser.CallStmtContext context)
            {
                if (context.CALL() != null)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
