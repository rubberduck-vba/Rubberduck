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
    public sealed class OptionExplicitInspection : InspectionBase
    {
        public OptionExplicitInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.ProceduralModule,
            DeclarationType.ClassModule
        };

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var results = UserDeclarations.ToList();

            var options = results
                .Where(declaration => declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionExplicitStmtContext)
                .ToList();

            var modules = results
                .Where(declaration => ModuleTypes.Contains(declaration.DeclarationType));

            var issues = modules.Where(module => !options.Select(option => option.Scope).Contains(module.Scope))
                .Select(issue => new OptionExplicitInspectionResult(this, issue));

            return issues;
        }
    }
}
