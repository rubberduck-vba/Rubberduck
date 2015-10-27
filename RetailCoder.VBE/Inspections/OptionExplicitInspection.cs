using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class OptionExplicitInspection : IInspection
    {
        public OptionExplicitInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "OptionExplicitInspection"; } }
        public string Description { get { return RubberduckUI.OptionExplicit; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.Module,
            DeclarationType.Class
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState parseResult)
        {
            var results = parseResult.Declarations().ToList();

            var options = results
                .Where(declaration => !declaration.IsBuiltIn 
                                      && declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionExplicitStmtContext)
                .ToList();

            var modules = results
                .Where(declaration => !declaration.IsBuiltIn && ModuleTypes.Contains(declaration.DeclarationType));

            var issues = modules.Where(module => !options.Select(option => option.Scope).Contains(module.Scope))
                .Select(issue => new OptionExplicitInspectionResult(this, issue.QualifiedName.QualifiedModuleName));

            return issues;
        }
    }
}