using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class EmptyStringLiteralInspection : InspectionBase
    {
        public EmptyStringLiteralInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.EmptyStringLiteralInspectionMeta; } }
        public override string Description { get { return InspectionsUI.EmptyStringLiteralInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {   
            return State.EmptyStringLiterals.Select(
                    context => new EmptyStringLiteralInspectionResult(this,
                            new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }
    }
}
