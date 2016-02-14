using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteLetStatementInspection : InspectionBase
    {
        public ObsoleteLetStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ObsoleteLetStatementInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteLetStatementInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            return State.ObsoleteLetContexts.Select(context =>
                new ObsoleteLetStatementUsageInspectionResult(this, new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }
    }
}