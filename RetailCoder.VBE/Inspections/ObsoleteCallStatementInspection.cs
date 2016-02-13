using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteCallStatementInspection : InspectionBase
    {
        public ObsoleteCallStatementInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public override string Description { get { return RubberduckUI.ObsoleteCall; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            return State.ObsoleteCallContexts.Select(context => 
                new ObsoleteCallStatementUsageInspectionResult(this,
                    new QualifiedContext<VBAParser.ExplicitCallStmtContext>(context.ModuleName, context.Context as VBAParser.ExplicitCallStmtContext)));
        }
    }
}