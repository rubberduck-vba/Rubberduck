using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class FunctionNotUsedInspection : IInspection
    {
        public FunctionNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.FunctionNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = parseResult.Declarations.Items.Where(declaration =>
                (declaration.DeclarationType == DeclarationType.Function)
                && !declaration.References.Any());

            foreach (var issue in declarations)
            {
                yield return new IdentifierNotUsedInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName.QualifiedModuleName);
            }
        }
    }
}