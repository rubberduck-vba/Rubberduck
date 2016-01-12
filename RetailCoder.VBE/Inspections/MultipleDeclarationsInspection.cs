using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class MultipleDeclarationsInspection : InspectionBase
    {
        public MultipleDeclarationsInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return RubberduckUI.MultipleDeclarations; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations
                .Where(item => !item.IsInspectionDisabled(AnnotationName))
                .Where(item => item.DeclarationType == DeclarationType.Variable
                            || item.DeclarationType == DeclarationType.Constant)
                .GroupBy(variable => variable.Context.Parent as ParserRuleContext)
                .Where(grouping => grouping.Count() > 1)
                .Select(grouping => new MultipleDeclarationsInspectionResult(this, Description, new QualifiedContext<ParserRuleContext>(grouping.First().QualifiedName.QualifiedModuleName, grouping.Key)));

            return issues;
        }
    }
}