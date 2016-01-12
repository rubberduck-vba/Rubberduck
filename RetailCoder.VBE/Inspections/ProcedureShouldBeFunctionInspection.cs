using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ProcedureShouldBeFunctionInspection : InspectionBase
    {
        public ProcedureShouldBeFunctionInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return InspectionsUI.ProcedureShouldBeFunctionInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            return State.ArgListsWithOneByRefParam
                .Where(context => context.Context.Parent is VBAParser.SubStmtContext)
                .Select(context => new ProcedureShouldBeFunctionInspectionResult(this,
                    State,
                    new QualifiedContext<VBAParser.ArgListContext>(context.ModuleName,
                        context.Context as VBAParser.ArgListContext),
                    new QualifiedContext<VBAParser.SubStmtContext>(context.ModuleName,
                        context.Context.Parent as VBAParser.SubStmtContext)));
        }
    }
}
