using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MultilineParameterInspection : InspectionBase
    {
        public MultilineParameterInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var multilineParameters = from p in UserDeclarations
                .Where(item => item.DeclarationType == DeclarationType.Parameter)
                where p.Context.GetSelection().LineCount > 1
                select p;

            var issues = multilineParameters
                .Select(param => new MultilineParameterInspectionResult(this, param));

            return issues;
        }
    }
}
