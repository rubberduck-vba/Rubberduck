using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    class GenericProjectNameInspection : IInspection
    {
        public GenericProjectNameInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "GenericProjectNameInspection"; } }
        public string Description { get { return RubberduckUI.GenericProjectName_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.Declarations.Items
                            .Where(declaration => !declaration.IsBuiltIn 
                                                && declaration.DeclarationType == DeclarationType.Project
                                                && declaration.IdentifierName.StartsWith("VBAProject"))
                            .Select(issue => new GenericProjectNameInspectionResult(string.Format(Description, issue.IdentifierName), Severity, issue, parseResult))
                            .ToList();

            return issues;
        }
    }
}
