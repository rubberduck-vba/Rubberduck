using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteCallStatementInspection : InspectionBase
    {
        public ObsoleteCallStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ObsoleteCallStatementInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteCallStatementInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            return State.ObsoleteCallContexts.Select(context => 
                new ObsoleteCallStatementUsageInspectionResult(this,
                    new QualifiedContext<VBAParser.ExplicitCallStmtContext>(context.ModuleName, context.Context as VBAParser.ExplicitCallStmtContext)));
        }
    }
}