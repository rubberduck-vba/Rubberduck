using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class MultipleDeclarationsInspection : IInspection
    {
        public MultipleDeclarationsInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "MultipleDeclarationsInspection"; } }
        public string Description { get { return RubberduckUI.MultipleDeclarations; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.Declarations.Items
                .Where(item => !item.IsBuiltIn)
                .Where(item => item.DeclarationType == DeclarationType.Variable
                            || item.DeclarationType == DeclarationType.Constant)
                .GroupBy(variable => variable.Context.Parent as ParserRuleContext)
                .Where(grouping => grouping.Count() > 1)
                .Select(grouping => new MultipleDeclarationsInspectionResult(Description, Severity, new QualifiedContext<ParserRuleContext>(grouping.First().QualifiedName.QualifiedModuleName, grouping.Key)));

            return issues;
        }
    }
}