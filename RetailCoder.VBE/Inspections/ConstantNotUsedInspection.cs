using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ConstantNotUsedInspection : IInspection
    {
        public ConstantNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.ConstantNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = parseResult.Declarations.Items.Where(declaration =>
                (declaration.DeclarationType == DeclarationType.Constant)
                && !declaration.References.Any());

            foreach (var issue in declarations)
            {
                yield return new IdentifierNotUsedInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName.QualifiedModuleName);
            }
        }
    }
}