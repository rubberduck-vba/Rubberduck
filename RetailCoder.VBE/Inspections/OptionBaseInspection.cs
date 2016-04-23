using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class OptionBaseInspection : InspectionBase
    {
        public OptionBaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.OptionBaseInspectionMeta; } }
        public override string Description { get { return InspectionsUI.OptionBaseInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var options = UserDeclarations
                .Where(declaration => declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionBaseStmtContext)
                .ToList();

            if (!options.Any())
            {
                return new List<InspectionResultBase>();
            }

            var issues = options.Where(option => ((VBAParser.OptionBaseStmtContext)option.Context).numberLiteral().GetText() == "1")
                                .Select(issue => new OptionBaseInspectionResult(this, issue.QualifiedName.QualifiedModuleName));

            return issues;
        }
    }
}