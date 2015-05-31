using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections
{
    /*class GenericProjectNameInspection : IInspection
    {
        public GenericProjectNameInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return RubberduckUI.GenericProjectName_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.Declarations.Items
                            .Where(declaration => declaration.DeclarationType == DeclarationType.Project
                                               && declaration.IdentifierName.Contains("VBAProject"))
                            .Select(issue => new GenericProjectNameInspectionResult(Name, Severity, issue.QualifiedName.QualifiedModuleName))
                            .ToList();

            return issues;
        }
    }*/
}
