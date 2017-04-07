using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MultipleDeclarationsInspection : InspectionBase
    {
        public MultipleDeclarationsInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var issues = UserDeclarations
                .Where(item => item.DeclarationType == DeclarationType.Variable
                            || item.DeclarationType == DeclarationType.Constant)
                .GroupBy(variable => variable.Context.Parent as ParserRuleContext)
                .Where(grouping => grouping.Count() > 1)
                .Select(grouping => new MultipleDeclarationsInspectionResult(this, new QualifiedContext<ParserRuleContext>(grouping.First().QualifiedName.QualifiedModuleName, grouping.Key)));

            return issues;
        }
    }
}
