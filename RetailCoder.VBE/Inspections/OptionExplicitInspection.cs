using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class OptionExplicitInspection : IInspection
    {
        private readonly IRubberduckFactory<IRubberduckCodePane> _factory;

        public OptionExplicitInspection(IRubberduckFactory<IRubberduckCodePane> factory)
        {
            _factory = factory;
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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var options = parseResult.Declarations.Items
                .Where(declaration => !declaration.IsBuiltIn 
                                      && declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionExplicitStmtContext)
                .ToList();

            var modules = parseResult.Declarations.Items
                .Where(declaration => !declaration.IsBuiltIn && ModuleTypes.Contains(declaration.DeclarationType));

            var issues = modules.Where(module => !options.Select(option => option.Scope).Contains(module.Scope))
                .Select(issue => new OptionExplicitInspectionResult(Description, Severity, issue.QualifiedName.QualifiedModuleName, _factory));

            return issues;
        }
    }
}