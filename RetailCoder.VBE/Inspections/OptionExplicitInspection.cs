using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class OptionExplicitInspection : InspectionBase
    {
        public OptionExplicitInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.OptionExplicitInspectionMeta; } }
        public override string Description { get { return InspectionsUI.OptionExplicitInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.ProceduralModule,
            DeclarationType.ClassModule
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
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
