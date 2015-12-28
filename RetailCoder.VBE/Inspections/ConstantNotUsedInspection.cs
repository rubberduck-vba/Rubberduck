using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var results = state.AllDeclarations.Where(declaration =>
                !declaration.IsBuiltIn 
                && declaration.DeclarationType == DeclarationType.Constant
                && !declaration.References.Any());

            return results.Select(issue => 
                new IdentifierNotUsedInspectionResult(this, issue, ((dynamic)issue.Context).ambiguousIdentifier(), issue.QualifiedName.QualifiedModuleName)).Cast<CodeInspectionResultBase>();
        }
    }
}