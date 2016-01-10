using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class OptionBaseInspection : InspectionBase
    {
        public OptionBaseInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return RubberduckUI.OptionBase; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var options = UserDeclarations
                .Where(declaration => declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionBaseStmtContext)
                .ToList();

            if (!options.Any())
            {
                return new List<CodeInspectionResultBase>();
            }

            var issues = options.Where(option => ((VBAParser.OptionBaseStmtContext)option.Context).INTEGERLITERAL().GetText() == "1")
                                .Select(issue => new OptionBaseInspectionResult(this, issue.QualifiedName.QualifiedModuleName));

            return issues;
        }
    }
}