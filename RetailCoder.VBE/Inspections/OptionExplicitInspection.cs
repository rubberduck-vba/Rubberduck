using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class OptionExplicitInspection : InspectionBase
    {
        public OptionExplicitInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return RubberduckUI.OptionExplicit; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.Module,
            DeclarationType.Class
        };

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var results = UserDeclarations.ToList();

            var options = results
                .Where(declaration => declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionExplicitStmtContext)
                .ToList();

            var modules = results
                .Where(declaration => ModuleTypes.Contains(declaration.DeclarationType));

            var issues = modules.Where(module => !options.Select(option => option.Scope).Contains(module.Scope))
                .Select(issue => new OptionExplicitInspectionResult(this, issue.QualifiedName.QualifiedModuleName));

            return issues;
        }
    }
}