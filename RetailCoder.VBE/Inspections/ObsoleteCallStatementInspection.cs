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
                    new QualifiedContext<VBAParser.ExplicitCallStmtContext>(context.ModuleName, context.Context as VBAParser.ExplicitCallStmtContext)));
        }

        public class ObsoleteCallStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.ExplicitCallStmtContext> _contexts = new List<VBAParser.ExplicitCallStmtContext>();
            public IEnumerable<VBAParser.ExplicitCallStmtContext> Contexts { get { return _contexts; } }

            public override void ExitExplicitCallStmt(VBAParser.ExplicitCallStmtContext context)
            {
                _contexts.Add(context);
            }
        }

    }
}