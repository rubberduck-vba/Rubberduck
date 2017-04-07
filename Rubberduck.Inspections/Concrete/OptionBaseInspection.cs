using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class OptionBaseInspection : InspectionBase
    {
        public OptionBaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var options = UserDeclarations
                .Where(declaration => declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionBaseStmtContext)
                .ToList();

            if (!options.Any())
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var issues = options.Where(option => ((VBAParser.OptionBaseStmtContext)option.Context).numberLiteral().GetText() == "1")
                                .Select(issue => new OptionBaseInspectionResult(this, issue));

            return issues;
        }
    }
}
