using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ConstantNotUsedInspection : IInspection
    {
        public ConstantNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ConstantNotUsedInspection"; } }
        public string Description { get { return RubberduckUI.ConstantNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = parseResult.Declarations.Items.Where(declaration =>
                !declaration.IsBuiltIn &&
                (declaration.DeclarationType == DeclarationType.Constant)
                && !declaration.References.Any());

            foreach (var issue in declarations)
            {
                yield return new IdentifierNotUsedInspectionResult(string.Format(Description, issue.IdentifierName), Severity, ((dynamic)issue.Context).ambiguousIdentifier(), issue.QualifiedName.QualifiedModuleName);
            }
        }
    }
}